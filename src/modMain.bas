Attribute VB_Name = "modMain"
'[modMain.bas]

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
' O7 IPSec / TroubleShooting / Certificates by Alex Dragokas
' O17 Policy Scripts, new keys, DHCP DNS by Alex Dragokas
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
'O7 - Policies: Regedit block / IPSec; O7 - TroubleShooting: system settings, that lead to OS malfunction
'O8 - IE Context menu item
'O9 - IE Tools menu item/button
'O10 - Winsock hijack
'O11 - IE Advanced Options group
'O12 - IE Plugin
'O13 - IE DefaultPrefix hijack
'O14 - IERESET.INF hijack
'O15 - Trusted Zone autoadd
'O16 - Downloaded Program Files
'O17 - Domain hijacks / DHCP DNS
'O18 - Protocol & Filter enum & Ports
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
' - increase LAST_CHECK_OTHER_SECTION_NUMBER const
' - add translation strings: after # 31, 261, 435

'Next possible methods:
'* SearchAccurates 'URL' method in a InitPropertyBag (??)
'* HKLM\..\CurrentVersion\ModuleUsage
'* HKLM\..\Internet Explorer\SafeSites (searchaccurate)

'New command line keys:
'
'/noBackup - отключает создание резервных копий во время фикса
'/install /autostart d:X - установка в планировщик заданий с запуском по логину юзера с задержкой в X сек.
'/instDir:"PATH" - PATH: путь к папке, куда производится установка (по умолчанию: "%ProgramFiles(x86)%\HiJackThis Fork").
'/tool:Autoruns - запуск проверки через SysInternals Autoruns
'/tool:Executed - запуск проверки через NirSoft ExecutedProgramList
'/tool:LastActivity - запуск проверки через NirSoft LastActivityView (только *.exe и *.dll файлы)
'/tool:ServiWin - запуск проверки через NirSoft ServiWin
'/tool:TaskScheduler - запуск проверки через NirSoft TaskSchedulerView
'/sigcheck - фильтрация по цифровой подписи Microsoft
'/vtcheck - выполнение проверки файлов на VT (через AutoRuns)
'/autofix:vt - выполнить удаление файлов с ненулевым детектом VT
'/delmode:pending - удаление файлов только в отложенном режиме (через перезагрузку)
'/delmode:disable - исправление элементов только через отключение (файлы не будут удалены, если это возможно)
'/reboot - перезагрузить систему, если были запланированы файлы на удаление
'/addfirewall - добавить AutoRuns в исключения брандмауера
'/fixHosts - выполнить очистку Hosts и сброс кеша сопоставителя (перед началом отгрузки на VT)
'/FixO4 - удаление элементов автозапуска (ключи реестра Run* и папка "Автозагрузка")
'/FixPolicy - сброс политик TaskMgr, Regedit, Explorer, TaskBar
'/FixCert - снятие блокировок ПО через сертификаты ЭЦП
'/FixIpSec - удаление политик IP Security (блокировки IP / открытые дыры в портах)
'/FixEnvVar - исправление некорректных настроек переменных окружения (%PATH%, %TEMP%)
'/FixO20 - сброс ключей WinLogon и App Init DLLs
'/FixTasks - удаление заданий планировщика
'/FixServices - удаление служб
'/FixWMIJob - удаление заданий WMI
'/FixIFEO - удаление ссылок на отладчики процессов, блокирующих запуск ПО
'/Disinfect - выполнить все доступные выше фиксы
'/FreezeProcess - заморозить все сторонние процессы перед выполнением фиксов
'/LockPoints - заблокировать точки автозапуска (остановка служб WMI, tasks, блокировка прав на запись ключей и файлов в папки автозапуска)
'/rawIgnoreList - активировать список игнорирования (белый список) в виде простого списка файлов .\whitelists.txt
'/Area:None - отключает выполнение основного сканирования HiJackThis.
'/noShortcuts - отключить создание ярлыков при установке через /install
'/! - останавливает парсинг ключей. Все, что находится после: используется в качестве ключей при установке HJT в автозапуск (планировщик заданий) вместо дефолтового /startupscan.


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
    Dim REG_REDIRECTED, REG_NOTREDIRECTED, REG_REDIRECTION_BOTH, _REG_REDIRECTION_NOT_DEFINED
#End If

Public Enum ENUM_REG_VALUE_TYPE_RESTORE
    REG_RESTORE_SAME = -1&
    REG_RESTORE_SZ = 1&
    REG_RESTORE_EXPAND_SZ = 2&
    REG_RESTORE_BINARY = 3&
    REG_RESTORE_DWORD = 4&
    'REG_RESTORE_LINK = 6&
    REG_RESTORE_MULTI_SZ = 7&
    REG_RESTORE_QWORD = 8&
End Enum
#If False Then
    Dim REG_RESTORE_SAME, REG_RESTORE_SZ, REG_RESTORE_EXPAND_SZ, REG_RESTORE_DWORD, REG_RESTORE_MULTI_SZ, REG_RESTORE_QWORD
#End If

Public Enum ENUM_CURE_BASED
    FILE_BASED = 1          ' if need to cure .File()
    REGISTRY_BASED = 2      ' if need to cure .Reg()
    INI_BASED = 4           ' if need to cure ini-file in .reg()
    PROCESS_BASED = 8       ' if need to kill/freeze a process
    SERVICE_BASED = 16      ' if need to delete/restore service .ServiceName
    CUSTOM_BASED = 32       ' individual rule, based on .Custom() settings
End Enum

#If False Then
    Dim FILE_BASED, REGISTRY_BASED, INI_BASED, PROCESS_BASED, SERVICE_BASED, CUSTOM_BASED
#End If

Public Enum ENUM_COMMON_ACTION_BASED
    USE_FEATURE_DISABLE = &H10000
End Enum

Public Enum ENUM_REG_ACTION_BASED
    REMOVE_KEY = 1&
    REMOVE_VALUE = 2&
    RESTORE_VALUE = 4&
    RESTORE_VALUE_INI = 8&
    REMOVE_VALUE_INI = &H10&
    REPLACE_VALUE = &H20&
    APPEND_VALUE_NO_DOUBLE = &H40&
    REMOVE_VALUE_IF_EMPTY = &H80&
    REMOVE_KEY_IF_NO_VALUES = &H100&
    TRIM_VALUE = &H200&
    BACKUP_KEY = &H400&
    BACKUP_VALUE = &H800&
    JUMP_KEY = &H1000&
    JUMP_VALUE = &H2000&
    RESTORE_KEY_PERMISSIONS = &H4000&
    RESTORE_KEY_PERMISSIONS_RECURSE = &H8000&
    USE_FEATURE_DISABLE_REG = &H10000
End Enum
#If False Then
    Dim REMOVE_KEY, REMOVE_VALUE, RESTORE_VALUE, RESTORE_VALUE_INI, REMOVE_VALUE_INI, REPLACE_VALUE
    Dim APPEND_VALUE_NO_DOUBLE, REMOVE_VALUE_IF_EMPTY, REMOVE_KEY_IF_NO_VALUES, TRIM_VALUE, BACKUP_KEY, BACKUP_VALUE
    Dim JUMP_KEY, JUMP_VALUE, RESTORE_KEY_PERMISSIONS, RESTORE_KEY_PERMISSIONS_RECURSE
#End If

Public Enum ENUM_FILE_ACTION_BASED
    REMOVE_FILE = 1
    REMOVE_FOLDER = 2
    RESTORE_FILE = 4   'not used yet
    RESTORE_FILE_SFC = 8
    UNREG_DLL = 16
    BACKUP_FILE = 32
    JUMP_FILE = 64
    JUMP_FOLDER = 128
    CREATE_FOLDER = 256
    USE_FEATURE_DISABLE_FILE = &H10000
End Enum
#If False Then
    Dim REMOVE_FILE, REMOVE_FOLDER, RESTORE_FILE, RESTORE_FILE_SFC, UNREG_DLL, BACKUP_FILE, JUMP_FILE, JUMP_FOLDER, CREATE_FOLDER
#End If

Public Enum ENUM_PROCESS_ACTION_BASED
    KILL_PROCESS = 1
    FREEZE_PROCESS = 2
    FREEZE_OR_KILL_PROCESS = 4
    USE_FEATURE_DISABLE_PROCESS = &H10000
End Enum
#If False Then
    Dim KILL_PROCESS, FREEZE_PROCESS, FREEZE_OR_KILL_PROCESS
#End If

Public Enum ENUM_SERVICE_ACTION_BASED
    DELETE_SERVICE = 1
    RESTORE_SERVICE = 2 ' not yet implemented
    DISABLE_SERVICE = 4
    ENABLE_SERVICE = 8
    USE_FEATURE_DISABLE_SERVICE = &H10000
End Enum
#If False Then
    Dim DELETE_SERVICE, RESTORE_SERVICE
#End If

Public Enum ENUM_CUSTOM_ACTION_BASED
    CUSTOM_ACTION_O25 = 1
    CUSTOM_ACTION_SPECIFIC = 2
End Enum
#If False Then
    Dim CUSTOM_ACTION_O25, CUSTOM_ACTION_SPECIFIC
#End If

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
    DateM           As Date
    SD              As String
End Type

Public Type FIX_FILE
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
    RunState        As SERVICE_STATE
    ActionType      As ENUM_SERVICE_ACTION_BASED
End Type

Private Type FIX_CUSTOM
    ActionType      As ENUM_CUSTOM_ACTION_BASED
    ObjectName      As String
End Type

Public Enum JUMP_ENTRY_TYPE
    JUMP_ENTRY_FILE = 1
    JUMP_ENTRY_REGISTRY = 2
End Enum

Private Type JUMP_ENTRY
    File()          As FIX_FILE
    Registry()      As FIX_REG_KEY
    Type            As JUMP_ENTRY_TYPE
End Type

Private Type O25_ActiveScriptConsumer_Entry
    File      As String
    Text      As String
    Engine    As String
End Type

Private Type O25_CommandLineConsumer_Entry
    ExecPath        As String
    WorkDir         As String
    CommandLine     As String
    Interactive     As Boolean
End Type

Public Enum O25_TIMER_TYPE
    O25_TIMER_ABSOLUTE = 1
    O25_TIMER_INTERVAL
End Enum

Private Type O25_Timer_Entry
    Type            As O25_TIMER_TYPE
    className       As String
    ID              As String
    Interval        As Long 'for O25_TIMER_INTERVAL
    EventDateTime   As Date 'for O25_TIMER_ABSOLUTE
End Type

Public Enum O25_CONSUMER_TYPE
    O25_CONSUMER_ACTIVE_SCRIPT = 1
    O25_CONSUMER_COMMAND_LINE
End Enum

Private Type O25_Consumer_Entry
    Name        As String
    NameSpace   As String
    Path        As String
    Type        As O25_CONSUMER_TYPE
    Script      As O25_ActiveScriptConsumer_Entry
    Cmd         As O25_CommandLineConsumer_Entry
    KillTimeout As Long
End Type

Private Type O25_Filter_Entry
    Name      As String
    NameSpace As String
    Path      As String
    Query     As String
End Type

Public Type O25_ENTRY
    Filter      As O25_Filter_Entry
    Timer       As O25_Timer_Entry
    Consumer    As O25_Consumer_Entry
End Type

Public Enum FIX_ITEM_STATE
    ITEM_STATE_DISABLED
    ITEM_STATE_ENABLED
End Enum

Public Type SCAN_RESULT
    HitLineW        As String
    HitLineA        As String
    Section         As String
    Alias           As String
    Name            As String
    State           As FIX_ITEM_STATE
    Reg()           As FIX_REG_KEY
    File()          As FIX_FILE
    Process()       As FIX_PROCESS
    Service()       As FIX_SERVICE
    Custom()        As FIX_CUSTOM
    Jump()          As JUMP_ENTRY
    CureType        As ENUM_CURE_BASED
    O25             As O25_ENTRY
    NoNeedBackup    As Boolean          'if no backup required / or impossible
    Reboot          As Boolean
    ForceMicrosoft  As Boolean
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
    dSafeProtocols As clsTrickHashTable
    dSafeFilters As clsTrickHashTable
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

Private Type FONT_PROPERTY
    Bold        As Boolean
    Italic      As Boolean
    Underline   As Boolean
    Size        As Long
End Type

Private Declare Sub OutputDebugStringA Lib "kernel32.dll" (ByVal lpOutputString As String)

Private HitSorted()     As String

Public gProcess()           As MY_PROC_ENTRY
Public g_TasksWL()          As TASK_WHITELIST_ENTRY
Public oDict                As DICTIONARIES

Public oDictFileExist       As clsTrickHashTable
Private dFontDefault        As clsTrickHashTable
Private aFontDefProp()      As FONT_PROPERTY

Public Scan()   As SCAN_RESULT    '// Dragokas. Used instead of parsing lines from result screen (like it was in original HJT 2.0.5).
                                  '// User type structures of arrays is filled together with using of method frmMain.lstResults.AddItem
                                  '// It is much efficiently and have Unicode support (native vb6 ListBox is ANSI only).
                                  '// Result screen will be replaced with CommonControls unicode aware controls by Krool (vbforums.com) in the nearest update,
                                  '// as well as StartupList2 by Merijn that currently use separate Microsoft MSCOMCTL.OCX library file.

Public Perf     As TYPE_PERFORMANCE

Public OSver    As clsOSInfo
Public Proc     As clsProcess
Public cMath    As clsMath

Private oDictSRV As clsTrickHashTable

Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal pszStrPtr As Long, ByVal Length As Long) As String


'it map ANSI scan result string from ListBox to Unicode string that is stored in memory (SCAN_RESULT structure)
Public Function GetScanResults(HitLineA As String, result As SCAN_RESULT, Optional out_idx As Long) As Boolean
    Dim i As Long
    For i = 1 To UBound(Scan)
        If HitLineA = Scan(i).HitLineA Then
            result = Scan(i)
            out_idx = i
            GetScanResults = True
            Exit Function
        End If
    Next
    'Cannot find appropriate cure item for:, "Error"
    MsgBoxW Translate(592) & vbCrLf & HitLineA, vbCritical, Translate(591)
End Function

Public Function RemoveFromScanResults(HitLineA As String) As Boolean
    Dim i As Long, j As Long
    Dim result As SCAN_RESULT
    For i = 1 To UBound(Scan)
        If HitLineA = Scan(i).HitLineA Then
            For j = i + 1 To UBound(Scan)
                Scan(j - 1) = Scan(j)
            Next
            If UBound(Scan) > 0 Then
                ReDim Preserve Scan(UBound(Scan) - 1)
            Else
                Scan(0) = result
            End If
            Exit For
        End If
    Next
End Function

' it add Unicode SCAN_RESULT structure to shared array
Public Sub AddToScanResults(result As SCAN_RESULT, Optional ByVal DoNotAddToListBox As Boolean, Optional DontClearResults As Boolean)
    Dim bFirstWarning As Boolean
    
    Const SelLastAdded As Boolean = False
    
    'LockWindowUpdate frmMain.lstResults.hwnd
    
    If bAutoLogSilent And Not g_bFixArg Then
        DoNotAddToListBox = True
    Else
        DoEvents
    End If
    If Not DoNotAddToListBox Then
        'checking if one of sections planned to be contains more then 50 entries -> block such attempt
        If Not SectionOutOfLimit(result.Section, bFirstWarning) Then
            frmMain.lstResults.AddItem result.HitLineW
            'select the last added line
            If SelLastAdded Then
                frmMain.lstResults.ListIndex = frmMain.lstResults.ListCount - 1
            End If
        Else
            If bFirstWarning Then
                frmMain.lstResults.AddItem result.Section & " - Too many entries ( > 250 )" '=> look Const LIMIT
                If SelLastAdded Then
                    frmMain.lstResults.ListIndex = frmMain.lstResults.ListCount - 1
                End If
            End If
        End If
    End If
    ReDim Preserve Scan(UBound(Scan) + 1)
    'Unicode to ANSI mapping (dirty hack)
    result.HitLineA = frmMain.lstResults.List(frmMain.lstResults.ListCount - 1)
    Scan(UBound(Scan)) = result
    'Sleep 5
    'LockWindowUpdate False
    'Erase Result struct
    If Not DontClearResults Then
        EraseScanResults result
    End If
End Sub

Public Sub EraseScanResults(result As SCAN_RESULT)
    Dim EmptyResult As SCAN_RESULT
    result = EmptyResult
End Sub

'// Increase number of sections +1 and returns TRUE, if total number > LIMIT
Private Function SectionOutOfLimit(p_Section As String, Optional bFirstWarning As Boolean, Optional bErase As Boolean) As Long
    Dim LIMIT As Long: LIMIT = 250&
    If p_Section = "O23" Then LIMIT = 750
    
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
    Dim result As SCAN_RESULT
    With result
        .Section = Section
        .HitLineW = HitLine
    End With
    AddToScanResults result, DoNotAddToListBox
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
        .Add "Software\Microsoft\Internet Explorer\Main,Default_Page_URL," & Default_Page_URL & "|http://www.msn.com|res://iesetup.dll/HardAdmin.htm|res://iesetup.dll/SoftAdmin.htm|res://shdoclc.dll/softAdmin.htm|,"
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
        .Add "Software\Microsoft\Internet Explorer\Main,Start Page,$DEFSTARTPAGE|http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=msnhome|http://www.microsoft.com/isapi/redir.dll?prd={SUB_PRD}&clcid={SUB_CLSID}&pver={SUB_PVER}&ar=home|res://iesetup.dll/HardAdmin.htm|res://iesetup.dll/SoftAdmin.htm|about:Tabs|about:NewsFeed|,"
        .Add "Software\Microsoft\Internet Explorer\Main,SearchURL,,"
        .Add "Software\Microsoft\Internet Explorer\Main,Start Page Redirect Cache,http://ru.msn.com/?ocid=iehp|,"
        
        .Add "Software\Microsoft\Internet Explorer\Search,SearchAssistant,$DEFSEARCHASS|,"
        .Add "Software\Microsoft\Internet Explorer\Search,CustomizeSearch,$DEFSEARCHCUST|,"
        .Add "Software\Microsoft\Internet Explorer\Search,(Default),,"
        
        .Add "Software\Microsoft\Internet Explorer\SearchURL,(Default),,"
        .Add "Software\Microsoft\Internet Explorer\SearchURL,SearchURL,,"
        
        .Add "Software\Microsoft\Internet Explorer\Main,Startpagina,,"
        .Add "Software\Microsoft\Internet Explorer\Main,First Home Page,|res://iesetup.dll/HardAdmin.htm|res://iesetup.dll/SoftAdmin.htm,"
        .Add "Software\Microsoft\Internet Explorer\Main,Local Page,%SystemRoot%\System32\blank.htm|%SystemRoot%\SysWOW64\blank.htm|%11%\blank.htm|,"
        .Add "Software\Microsoft\Internet Explorer\Main,Start Page_bak,,"
        .Add "Software\Microsoft\Internet Explorer\Main,HomeOldSP,,"
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
        
        .Add "Software\Microsoft\Internet Explorer\Toolbar,LinksFolderName,Links|" & STR_CONST.RU_LINKS & "|," 'Ссылки
        
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
        .Add "REG:system.ini;boot;UserInit;%WINDIR%\System32\userinit.exe|userinit.exe|userinit;"
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
    
    ' === LOAD SAFE O5 CONTROL PANEL DISABLED ITEMS
    
    ReDim aSafeO5Items(0)
    
    If OSver.IsWindows10OrGreater Then
        sSafeO5Items_HKLM = "appwiz.cpl|bthprops.cpl|desk.cpl|Firewall.cpl|hdwwiz.cpl|inetcpl.cpl|intl.cpl|irprops.cpl|joy.cpl|main.cpl|mmsys.cpl|ncpa.cpl|powercfg.cpl|sysdm.cpl|tabletpc.cpl|telephon.cpl|timedate.cpl"
        sSafeO5Items_HKLM_32 = "appwiz.cpl|bthprops.cpl|desk.cpl|Firewall.cpl|hdwwiz.cpl|inetcpl.cpl|intl.cpl|irprops.cpl|joy.cpl|main.cpl|mmsys.cpl|ncpa.cpl|powercfg.cpl|sysdm.cpl|telephon.cpl|timedate.cpl"
    ElseIf OSver.IsWindows8OrGreater Then
        sSafeO5Items_HKLM = "sysdm.cpl|inetcpl.cpl|ncpa.cpl|tabletpc.cpl|joy.cpl|powercfg.cpl|Firewall.cpl|telephon.cpl|irprops.cpl|intl.cpl|timedate.cpl|hdwwiz.cpl|mmsys.cpl|desk.cpl|main.cpl|appwiz.cpl|bthprops.cpl"
        sSafeO5Items_HKLM_32 = "sysdm.cpl|inetcpl.cpl|ncpa.cpl|Firewall.cpl|telephon.cpl|powercfg.cpl|irprops.cpl|joy.cpl|intl.cpl|timedate.cpl|hdwwiz.cpl|mmsys.cpl|main.cpl|desk.cpl|appwiz.cpl|bthprops.cpl"
    ElseIf (OSver.MajorMinor = 6.1 And OSver.IsServer) Then '2008 Server R2
        sSafeO5Items_HKLM = "hdwwiz.cpl|telephon.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|inetcpl.cpl|joy.cpl|mmsys.cpl|Firewall.cpl|powercfg.cpl|intl.cpl|timedate.cpl|main.cpl|tabletpc.cpl|infocardcpl.cpl"
        sSafeO5Items_HKLM_32 = "hdwwiz.cpl|telephon.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|inetcpl.cpl|joy.cpl|mmsys.cpl|Firewall.cpl|powercfg.cpl|intl.cpl|timedate.cpl|main.cpl|tabletpc.cpl|infocardcpl.cpl"
    ElseIf OSver.IsWindows7OrGreater Then
        sSafeO5Items_HKLM = "hdwwiz.cpl|telephon.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|inetcpl.cpl|joy.cpl|mmsys.cpl|Firewall.cpl|powercfg.cpl|intl.cpl|timedate.cpl|main.cpl|collab.cpl|irprops.cpl|tabletpc.cpl|infocardcpl.cpl|bthprops.cpl"
        sSafeO5Items_HKLM_32 = "hdwwiz.cpl|telephon.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|inetcpl.cpl|joy.cpl|mmsys.cpl|Firewall.cpl|powercfg.cpl|intl.cpl|timedate.cpl|main.cpl|collab.cpl|irprops.cpl|infocardcpl.cpl|bthprops.cpl"
    ElseIf OSver.IsWindowsVistaOrGreater Then
        sSafeO5Items_HKLM = "hdwwiz.cpl|appwiz.cpl|ncpa.cpl|sysdm.cpl|desk.cpl|Firewall.cpl|powercfg.cpl|infocardcpl.cpl|bthprops.cpl"
        sSafeO5Items_HKLM_32 = "ncpa.cpl|sysdm.cpl|desk.cpl|hdwwiz.cpl|Firewall.cpl|powercfg.cpl|appwiz.cpl|infocardcpl.cpl|bthprops.cpl"
    ElseIf OSver.IsWindowsXPOrGreater Then
        sSafeO5Items_HKU = "ncpa.cpl|odbccp32.cpl"
        sSafeO5Items_HKLM = "speech.cpl|infocardcpl.cpl"
    Else
        sSafeO5Items_HKU = "ncpa.cpl|odbccp32.cpl"
    End If
    
    
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
        .Add "http://go.microsoft.com"
        .Add "www.microsoft.com"
        .Add "microsoft.com"
        .Add "http://windowsupdate.com"
        .Add "http://runonce.msn.com"
        .Add "http://*.update.microsoft.com"
        .Add "https://*.update.microsoft.com"
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
    
    Set oDict.dSafeProtocols = New clsTrickHashTable
    oDict.dSafeProtocols.CompareMode = vbTextCompare

    '//TODO: O18 - add file path checking to database

    With oDict.dSafeProtocols
        .Add "ms-itss", "{0A9007C0-4076-11D3-8789-0000F8105754}"
        .Add "about", "{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "belarc", "{6318E0AB-2E93-11D1-B8ED-00608CC9A71F}"
        .Add "BPC", "{3A1096B3-9BFA-11D1-AE77-00C04FBBDEBC}"
        .Add "CDL", "{3DD53D40-7B8B-11D0-B013-00AA0059CE02}"
        .Add "cdo", "{CD00020A-8B95-11D1-82DB-00C04FB1625D}"
        .Add "copernicagentcache", "{AAC34CFD-274D-4A9D-B0DC-C74C05A67E1D}"
        .Add "copernicagent", "{A979B6BD-E40B-4A07-ABDD-A62C64A4EBF6}"
        .Add "dodots", "{9446C008-3810-11D4-901D-00B0D04158D2}"
        .Add "DVD", "{12D51199-0DB5-46FE-A120-47A3D7D937CC}"
        .Add "file", "{79EAC9E7-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ftp", "{79EAC9E3-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "gopher", "{79EAC9E4-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "https", "{79EAC9E5-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "http", "{79EAC9E2-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ic32pp", "{BBCA9F81-8F4F-11D2-90FF-0080C83D3571}"
        .Add "ipp", ""
        .Add "its", "{9D148291-B9C8-11D0-A4CC-0000F80149F6}"
        .Add "javascript", "{3050F3B2-98B5-11CF-BB82-00AA00BDCE0B}" '","<SysRoot>\System32\mshtml.dll
        .Add "junomsg", "{C4D10830-379D-11D4-9B2D-00C04F1579A5}"
        .Add "lid", "{5C135180-9973-46D9-ABF4-148267CBB8BF}"
        .Add "local", "{79EAC9E7-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "mailto", "{3050F3DA-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "mctp", "{D7B95390-B1C5-11D0-B111-0080C712FE82}"
        .Add "mhtml", "{05300401-BCBC-11D0-85E3-00C04FD85AB4}"
        .Add "mk", "{79EAC9E6-BAF9-11CE-8C82-00AA004BA90B}"
        .Add "ms-its50", "{F8606A00-F5CF-11D1-B6BB-0000F80149F6}"
        .Add "ms-its51", "{F6F1E82D-DE4D-11D2-875C-0000F8105754}"
        .Add "ms-its", "{9D148291-B9C8-11D0-A4CC-0000F80149F6}"
        .Add "mso-offdap", "{3D9F03FA-7A94-11D3-BE81-0050048385D1}"
        .Add "ndwiat", "{13F3EA8B-91D7-4F0A-AD76-D2853AC8BECE}"
        .Add "res", "{3050F3BC-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "sysimage", "{76E67A63-06E9-11D2-A840-006008059382}"
        .Add "tve-trigger", "{CBD30859-AF45-11D2-B6D6-00C04FBBDE6E}"
        .Add "tv", "{CBD30858-AF45-11D2-B6D6-00C04FBBDE6E}"
        .Add "vbscript", "{3050F3B2-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "vnd.ms.radio", "{3DA2AA3B-3D96-11D2-9BD2-204C4F4F5020}"
        .Add "wia", "{13F3EA8B-91D7-4F0A-AD76-D2853AC8BECE}"
        .Add "mso-offdap11", "{32505114-5902-49B2-880A-1F7738E5A384}"
        .Add "DirectDVD", "{85A81A02-336B-43FF-998B-FE8E194FBA4D}"
        .Add "pcn", "{D540F040-F3D9-11D0-95BE-00C04FD93CA5}"
        .Add "msencarta", "{74D92DF3-6D9D-11D1-8B38-006097DBED7A}"
        .Add "msero", "{B0D92A71-886B-453B-A649-1B91F93801E7}"
        .Add "msref", "{74D92DF3-6D9D-11D1-8B38-006097DBED7A}"
        .Add "df2", "{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df3", "{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df4", "{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df5", "{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df23chat", "{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "df5demo", "{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "ofpjoin", "{219A97F3-D661-4766-B658-646A771AE49E}"
        .Add "saphtmlp", "{D1F8BD1E-7967-11D2-B43A-006094B9EADB}"
        .Add "sapr3", "{D1F8BD1E-7967-11D2-B43A-006094B9EADB}"
        .Add "lbxfile", "{56831180-F115-11D2-B6AA-00104B2B9943}"
        .Add "lbxres", "{24508F1B-9E94-40EE-9759-9AF5795ADF52}"
        .Add "cetihpz", "{CF184AD3-CDCB-4168-A3F7-8E447D129300}"
        .Add "aim", "{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "shell", "{3050F406-98B5-11CF-BB82-00AA00BDCE0B}"
        .Add "asp", "{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "hsp", "{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-asp", "{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-hsp", "{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "x-zip", "{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "zip", "{8D32BA61-D15B-11D4-894B-000000000000}"
        .Add "bega", "{A57721C9-B905-49B3-8BCA-B99FBB8C627E}"
        .Add "bt2", "{1730B77B-F429-498F-9B15-4514D83C8294}"
        .Add "copernicdesktopsearch", "{D9656C75-5090-45C3-B27E-436FBC7ACFA7}"
        .Add "crick", "{B861500A-A326-11D3-A248-0080C8F7DE1E}"
        .Add "dadb", "{82D6F09F-4AC2-11D3-8BD9-0080ADB8683C}"
        .Add "dialux", "{8352FA4C-39C6-11D3-ADBA-00A0244FB1A2}"
        .Add "emistp", "{0EFAEA2E-11C9-11D3-88E3-0000E867A001}"
        .Add "ezstor", "{6344A3A0-96A7-11D4-88CC-000000000000}"
        .Add "flowto", "{C7101FB0-28FB-11D5-883A-204C4F4F5020}"
        .Add "g7ps", "{9EACF0FB-4FC7-436E-989B-3197142AD979}"
        .Add "intu-res", "{9CE7D474-16F9-4889-9BB9-53E2008EAE8A}"
        .Add "iwd", "{EA5F5649-A6C7-11D4-9E3C-0020AF0FFB56}"
        .Add "mavencache", "{DB47FDC2-8C38-4413-9C78-D1A68BF24EED}"
        .Add "ms-help", "{314111C7-A502-11D2-BBCA-00C04F8EC294}"
        .Add "msnim", "{828030A1-22C1-4009-854F-8E305202313F}"
        .Add "myrm", "{4D034FC3-013F-4B95-B544-44D49ABE3E76}"
        .Add "nbso", "{DF700763-3EAD-4B64-9626-22BEEFF3EA47}"
        .Add "nim", "{3D206AE2-3039-413B-B748-3ACC562EC22A}"
        .Add "OWC11.mso-offdap", "{32505114-5902-49B2-880A-1F7738E5A384}"
        .Add "pcl", "{182D0C85-206F-4103-B4FA-DCC1FB0A0A44}"
        .Add "pure-go", "{4746C79A-2042-4332-8650-48966E44ABA8}"
        .Add "qrev", "{9DE24BAC-FC3C-42C4-9FC4-76B3FAFDBD90}"
        .Add "rmh", "{23C585BB-48FF-4865-8934-185F0A7EB84C}"
        .Add "SafeAuthenticate", "{8125919B-9BE9-4213-A1D6-75188A22D21E}"
        .Add "sds", "{79E0F14C-9C52-4218-89A7-7C4B0563D121}"
        .Add "siteadvisor", "{3A5DC592-7723-4EAA-9EE6-AF4222BCF879}"
        .Add "smscrd", "{FA3F5003-93D4-11D2-8E48-00A0C98BD8C3}"
        .Add "stibo", "{FFAD3420-6D61-44F6-BA25-293F17152D79}"
        .Add "textwareilluminatorbase", "{CE5CD329-1650-414A-8DB0-4CBF72FAED87}"
        .Add "widimg", "{EE7C2AFF-5742-44FF-BD0E-E521B0D3C3BA}"
        .Add "wlmailhtml", "{03C514A3-1EFB-4856-9F99-10D7BE1653C0}"
        .Add "x-atng", "{7E8717B0-D862-11D5-8C9E-00010304F989}"
        .Add "x-excid", "{9D6CC632-1337-4A33-9214-2DA092E776F4}"
        .Add "x-mem1", "{C3719F83-7EF8-4BA0-89B0-3360C7AFB7CC}"
        .Add "x-mem3", "{4F6D06DD-44AB-4F89-BF13-9027B505B15A}"
        .Add "ct", "{774E529C-2458-48A2-8F57-3ED3105D8612}"
        .Add "cw", "{774E529C-2458-48A2-8F57-3ED3105D8612}"
        .Add "eti", "{3AAE7392-E7AA-11D2-969E-00105A088846}"
        .Add "livecall", "{828030A1-22C1-4009-854F-8E305202313F}"
        .Add "tbauth", "{14654CA6-5711-491D-B89A-58E571679951}"
        .Add "windows.tbauth", "{14654CA6-5711-491D-B89A-58E571679951}"
        .Add "msdaipp", "{E1D2BF40-A96B-11D1-9C6B-0000F875AC61}|{E1D2BF42-A96B-11D1-9C6B-0000F875AC61}"
    End With
        
    ' === LOAD FILTER SAFELIST === (O18)
    
    Set oDict.dSafeFilters = New clsTrickHashTable
    oDict.dSafeFilters.CompareMode = vbTextCompare

    With oDict.dSafeFilters
        .Add "application/octet-stream", "{1E66F26B-79EE-11D2-8710-00C04F79ED0D}|{F969FE8E-1937-45AD-AF42-8A4D11CBDC2A}"
        .Add "application/x-msdownload", "{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "application/vnd-backup-octet-stream", "{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "application/x-complus", "{1E66F26B-79EE-11D2-8710-00C04F79ED0D}"
        .Add "Class Install Handler", "{32B533BB-EDAE-11d0-BD5A-00AA00B92AF1}"
        .Add "deflate", "{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "gzip", "{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "lzdhtml", "{8f6b0360-b80d-11d0-a9b3-006097942311}"
        .Add "text/webviewhtml", "{733AC4CB-F1A4-11d0-B951-00A0C90312E1}"
        .Add "text/xml", "{807553E5-5146-11D5-A672-00B0D022E945}|{807563E5-5146-11D5-A672-00B0D022E945}|{32F66A26-7614-11D4-BD11-00104BD3F987}"
        .Add "application/x-icq", "{db40c160-09a1-11d3-baf2-000000000000}"
        .Add "application/msword", "{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd.ms-excel", "{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd.ms-powerpoint", "{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/x-microsoft-rpmsg-message", "{DFF82902-0B96-3B98-6F62-D655E146A23A}"
        .Add "application/vnd-viewer", "{CD4527E8-4FC7-48DB-9806-10537B501237}"
        .Add "application/x-bt2", "{6E1DDCE8-76BC-4390-9488-806E8FB1AD77}"
        .Add "application/x-internet-signup", "{A173B69A-1F9B-4823-9FDA-412F641E65D6}"
        .Add "text/html", "{8D42AD12-D7A1-4797-BCB7-AD89E5FCE4F7}|{F79B2338-A6E7-46D4-9201-422AA6E74F43}"
        .Add "text/x-mrml", "{C51721BE-858B-4A66-A8BF-D2882FF49820}"
        .Add "application/xhtml+xml", "{32F66A26-7614-11D4-BD11-00104BD3F987}"
    End With


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
    sSafeWinlogonNotify = "*crypt32chain*cryptnet*cscdll*ScCertProp*Schedule*SensLogn*termsrv*wlballoon*igfxcui*AtiExtEvent*wzcnotif*" & _
                          "ActiveSync*atmgrtok*avldr*Caveo*ckpNotify*Command AntiVirus Download*ComPlusSetup*CwWLEvent*dimsntfy*DPWLN*EFS*FolderGuard*GoToMyPC*IfxWlxEN*igfxcui*IntelWireless*klogon*LBTServ*LBTWlgn*LMIinit*loginkey*MCPClient*MetaFrame*NavLogon*NetIdentity Notification*nwprovau*OdysseyClient*OPXPGina*PCANotify*pcsinst*PFW*PixVue*ppeclt*PRISMAPI.DLL*PRISMGNA.DLL*psfus*QConGina*RAinit*RegCompact*SABWinLogon*SDNotify*Sebring*STOPzilla*sunotify*SymcEventMonitors*T3Notify*TabBtnWL*Timbuktu Pro*tpfnf2*tpgwlnotify*tphotkey*VESWinlogon*WB*WBSrv*WgaLogon*wintask*WLogon*WRNotifier*Zboard*zsnotify*sclgntfy*"
    
    sSafeIfeVerifier = "*vrfcore.dll*vfbasics.dll*vfcompat.dll*vfluapriv.dll*vfprint.dll*vfnet.dll*vfntlmless.dll*vfnws.dll*vfcuzz.dll*"
    
    'Loading Safe DNS list
    'https://www.comss.ru/list.php?c=securedns
    
    'These are checked with nslookup
    
    With colSafeDNS
        .Add "Google", "8.8.8.8"
        .Add "Google", "8.8.4.4"
        .Add "Verisign", "64.6.64.6"
        .Add "Verisign", "64.6.65.6"
        .Add "SkyDNS", "193.58.251.251"
        .Add "Cisco OpenDNS", "208.67.222.222"
        .Add "Cisco OpenDNS", "208.67.220.220"
        .Add "Cisco OpenDNS", "208.67.222.123"
        .Add "Cisco OpenDNS", "208.67.220.123"
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
        .Add "Adguard DNS", "176.103.130.132"
        .Add "Adguard DNS", "176.103.130.134"
        .Add "Adguard DNS", "176.103.130.136"
        .Add "Adguard DNS", "176.103.130.137"
        .Add "Yandex.DNS", "77.88.8.8"
        .Add "Yandex.DNS", "77.88.8.1"
        .Add "Yandex.DNS", "77.88.8.88"
        .Add "Yandex.DNS", "77.88.8.2"
        .Add "Yandex.DNS", "77.88.8.7"
        .Add "Yandex.DNS", "77.88.8.3"
        .Add "Comodo Secure DNS", "8.26.56.26"
        .Add "Comodo Secure DNS", "8.20.247.20"
        .Add "Comodo Secure DNS", "8.26.56.10"
        .Add "Comodo Secure DNS", "8.20.247.10"
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
        .Add "FreeDNS", "172.104.237.57"
        .Add "FreeDNS", "172.104.49.100"
        .Add "FreeDNS", "45.33.97.5"
        .Add "Alternate DNS", "198.101.242.72"
        .Add "Alternate DNS", "23.253.163.53"
        .Add "Rejector", "95.154.128.32"
        .Add "Rejector", "78.46.36.8"
        .Add "SmartViper", "208.76.50.50"
        .Add "SmartViper", "208.76.51.51"
        .Add "Neustar UltraDNS", "156.154.70.1"
        .Add "Neustar UltraDNS", "156.154.71.1"
        .Add "Neustar UltraDNS", "156.154.70.5"
        .Add "Neustar UltraDNS", "156.154.71.5"
        .Add "Neustar UltraDNS", "156.154.70.2"
        .Add "Neustar UltraDNS", "156.154.71.2"
        .Add "Neustar UltraDNS", "156.154.70.3"
        .Add "Neustar UltraDNS", "156.154.71.3"
        .Add "GreenTeamDNS", "81.218.119.11"
        .Add "GreenTeamDNS", "209.88.198.133"
        .Add "GTE", "192.76.85.133"
        .Add "GTE", "206.124.64.1"
        .Add "Hurricane Electric", "74.82.42.42"
        .Add "puntCAT", "109.69.8.51"
        .Add "Sprintlink General DNS", "204.117.214.10"
        .Add "Sprintlink General DNS", "199.2.252.10"
        .Add "Sprintlink General DNS", "204.97.212.10"
        .Add "Chaos Computer Club", "194.150.168.168"
        .Add "Chaos Computer Club", "213.73.91.35"
        .Add "Chaos Computer Club", "85.214.20.141"
        .Add "UncensoredDNS", "89.233.43.71"
        .Add "UncensoredDNS", "91.239.100.100"
        .Add "CyberGhost", "38.132.106.139"
        .Add "CyberGhost", "194.187.251.67"
        .Add "CyberGhost", "185.93.180.131"
        .Add "CyberGhost", "209.58.179.186"
        .Add "CyberGhost", "27.50.70.139"
        .Add "DNSReactor", "45.55.155.25"
        .Add "DNSReactor", "104.236.210.29"
        .Add "FDN", "80.67.169.12"
        .Add "FDN", "80.67.169.40"
        .Add "Lightning Wire Labs", "81.3.27.54"
        .Add "Lightning Wire Labs", "74.113.60.185"
        .Add "Freenom", "80.80.80.80"
        .Add "Freenom", "80.80.81.81"
        .Add "Quad9", "9.9.9.9"
        .Add "Quad9", "9.9.9.10"
        .Add "Quad9", "9.9.9.11"
        .Add "Quad9", "149.112.112.112"
        .Add "Quad9", "149.112.112.11"
        .Add "Quad9", "149.112.112.10"
        .Add "Xiala", "77.109.148.136"
        .Add "Xiala", "77.109.148.137"
        .Add "Cloudflare / APNIC", "1.1.1.1"
        .Add "Cloudflare / APNIC", "1.0.0.1"
        .Add "Cloudflare / APNIC", "1.1.1.2"
        .Add "Cloudflare / APNIC", "1.0.0.2"
        .Add "Cloudflare / APNIC", "1.1.1.3"
        .Add "Cloudflare / APNIC", "1.0.0.3"
        .Add "CleanBrowsing", "185.228.168.9"
        .Add "CleanBrowsing", "185.228.168.10"
        .Add "CleanBrowsing", "185.228.168.168"
        .Add "CleanBrowsing", "185.228.169.168"
        .Add "CleanBrowsing", "185.228.169.11"
        .Add "CleanBrowsing", "185.228.169.9"
        .Add "CenturyLink", "205.171.3.66"
        .Add "CenturyLink", "205.171.3.26"
        .Add "CenturyLink", "205.171.202.166"
        .Add "CenturyLink", "205.171.2.26"
        .Add "OpenNIC", "192.71.245.208"
        .Add "OpenNIC", "94.247.43.254"
        .Add "OpenNIC", "51.15.98.97"
        .Add "OpenNIC", "195.10.195.195"
        .Add "Fourth Estate", "45.77.165.194"
        .Add "Fourth Estate", "45.32.36.36"
        .Add "Safe Surfer", "104.197.28.121"
        .Add "Safe Surfer", "104.155.237.225"
        .Add "Comss.one", "92.38.152.163"
        .Add "Comss.one", "93.115.24.204"
        .Add "Comss.one", "92.223.109.31"
        .Add "Comss.one", "91.230.211.67"
    End With
    
    With colDisallowedCert
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
        'additional certificates on XP
        .Add "e-islem.kktcmerkezbankasi.org", "F92BE5266CC05DB2DC0DC3F2DC74E02DEFD949CB"
        .Add "Microsoft Online Svcs BPOS CA2", "F5A874F3987EB0A9961A564B669A9050F770308A"
        .Add "Microsoft Online Svcs BPOS EMEA CA2", "E9809E023B4512AA4D4D53F40569C313C1D0294D"
        .Add "Microsoft Online Svcs BPOS APAC CA3", "E95DD86F32C771F0341743EBD75EC33C74A3DED9"
        .Add "Microsoft Genuine Windows Phone Public Preview CA01", "E38A2B7663B86796436D8DF5898D9FAA6835B238"
        .Add "Microsoft Online Svcs BPOS APAC CA2", "D8CE8D07F9F19D2569C2FB854401BC99C1EB7C3B"
        .Add "Microsoft Online Svcs BPOS APAC CA1", "D43153C8C25F0041287987250F1E3CABAC8C2177"
        .Add "Microsoft Online Svcs BPOS APAC CA5", "D0BB3E3DFBFB86C0EEE2A047E328609E6E1F185E"
        .Add "*.EGO.GOV.TR", "C69F28C825139E65A646C434ACA5A1D200295DB1"
        .Add "Microsoft IPTVe CA", "BED412B1334D7DFCEBA3015E5F9F905D571C45CF"
        .Add "Microsoft Online Svcs CA5", "A81706D31E6F5C791CD9D3B1B9C63464954BA4F5"
        .Add "Microsoft Online Svcs BPOS EMEA CA3", "A7B5531DDC87129E2C3BB14767953D6745FB14A6"
        .Add "Microsoft Online Svcs BPOS EMEA CA1", "A35A8C727E88BCCA40A3F9679CE8CA00C26789FD"
        .Add "Microsoft Online Svcs CA1", "A221D360309B5C3C4097C44CC779ACC5A9845B66"
        .Add "Microsoft Online CA001", "A1505D9843C826DD67ED4EA5209804BDBB0DF502"
        .Add "Microsoft Online Svcs CA3", "8977E8569D2A633AF01D0394851681CE122683A6"
        .Add "Microsoft Online Svcs BPOS EMEA CA6", "838FFD509DE868F481C29819992E38A4F7082873"
        .Add "Microsoft Online Svcs BPOS CA1", "7613BF0BA261006CAC3ED2DDBEF343425357F18B"
        .Add "Microsoft Online Svcs CA4", "6690C02B922CBD3FF0D0A5994DBD336592887E3F"
        .Add "Microsoft Online Svcs CA4", "5D5185DF1EB7DC76015422EC8138A5724BEE2886"
        .Add "AC DG Tresor SSL", "5CE339465F41A1E423149F65544095404DE6EBE2"
        .Add "Microsoft Online Svcs BPOS CA2", "587B59FB52D8A683CBE1CA00E6393D7BB923BC92"
        .Add "Microsoft Online Svcs BPOS CA2", "4ED8AA06D1BC72CA64C47B1DFE05ACC8D51FC76F"
        .Add "Microsoft Online Svcs CA5", "4DF13947493CFF69CDE554881C5F114E97C3D03B"
        .Add "*.google.com", "4D8547B7F864132A7F62D9B75B068521F10B68E3"
        .Add "CN=Microsoft Online Svcs BPOS APAC CA4", "3A26012171855D4020C973BEC3F4F9DA45BD2B83"
        .Add "Microsoft Online Svcs CA3", "374D5B925B0BD83494E656EB8087127275DB83CE"
        .Add "Microsoft Online Svcs BPOS EMEA CA4", "330D8D3FD325A0E5FDDDA27013A2E75E7130165F"
        .Add "Microsoft Online Svcs CA1", "23EF3384E21F70F034C467D4CBA6EB61429F174E"
        .Add "Microsoft Online Svcs CA6", "09FF2CC86CEEFA8A8BB3F2E3E84D6DA3FABBF63E"
        .Add "Microsoft Online Svcs BPOS EMEA CA5", "09271DD621EBD3910C2EA1D059F99B8181405A17"
        .Add "Microsoft Online Svcs BPOS APAC CA6", "08738A96A4853A52ACEF23F782E8E1FEA7BCED02"
        'fresh update
        .Add "DSDTestProvider", "02C2D931062D7B1DC2A5C7F5F0685064081FB221"
        .Add "www.live.fi", "08E4987249BC450748A4A78133CBF041A3510033"
        .Add "D-LINK CORPORATION", "3EB44E5FFE6DC72DED703E99902722DB38FFD1CB"
        .Add "NIC Certifying Authority", "4822824ECE7ED1450C039AA077DC1F8AE3489BBF"
        .Add "Alpha Networks Inc.", "7311E77EC400109D6A5326D8F6696204FD59AA3B"
        .Add "*.xboxlive.com", "8B2E65A5DA17FCCCBCDE7EF87B0C0ED5D0701F9F"
        .Add "KEEBOX, INC", "915A478DB939925DA8D9AEA12D8BBA140D26599C"
        .Add "eDellRoot", "98A04E4163357790C4A79E6D713FF0AF51FE6927"
        .Add "NIC CA 2011", "C6796490CDEEAAB31AED798752ECD003E6866CB2"
        .Add "NIC CA 2014", "D2DBF71823B2B8E78F5958096150BFCB97CC388A"
        .Add "TRENDnet, Inc.", "DB5042ED256FF426867B332887ECCE2D95E79614"
        .Add "MCSHOLDING TEST", "E1F3591E769865C4E447ACC37EAFC9E2BFE4C576"
        'updated 30 may 2020
        .Add "DarkMatter High Assurance CA", "D3FD325D0F2259F693DD789430E3A9430BB59B98"
        .Add "127.0.0.1", "C597D4E7FF9CE5BD3EC321C11827FCA9294A6BA1"
        .Add "DarkMatter Assured CA", "9FEB091E053D1C453C789E8E9C446D31CB177ED9"
        .Add "DarkMatter High Assurance CA", "8835437D387BBB1B58FF5A0FF8D003D8FE04AED4"
        .Add "DarkMatter Assured CA", "6B6FA65B1BDC2A0F3A7E66B590F93297B8EB56B9"
        .Add "DarkMatter Secure CA", "6A2C691767C2F1999B8C020CBAB44756A99A0C41"
        .Add "DarkMatter Secure CA", "3AD010247A8F1E991F8DDE5D47989CB5202E5614"
        .Add "SenncomRootCA", "1990649205B55EAB5D692E9EDB1BE0DDD3B037DE"
        
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
        If g_hDebugLog = 0 Then OpenDebugLogHandle
    End If
    
    If Not bAutoLog Then Perf.StartTime = GetTickCount()
    
    bScanMode = True
    
    SetPriorityAllThreads GetCurrentProcess(), THREAD_PRIORITY_HIGHEST
    
    frmMain.txtNothing.ZOrder 1
    frmMain.txtNothing.Visible = False
    
    'frmMain.shpBackground.Tag = iItems
    SetProgressBar g_HJT_Items_Count   'R + F + O26
    
    If Not bAutoLogSilent Then
        Call GetProcesses(gProcess)
    Else
        If AryPtr(gProcess) = 0 Then
            Call GetProcesses(gProcess)
        End If
    End If
    
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
        If .lstResults.ListCount > 0 Or bAutoLogSilent Then
            .txtNothing.ZOrder 1
            .txtNothing.Visible = False
            '.cmdFix.Enabled = True
            '.cmdFix.FontBold = True
            '.cmdSaveDef.Enabled = True
        Else
            .txtNothing.Visible = True
            .txtNothing.ZOrder 0
            '.cmdFix.FontBold = False
            '.cmdFix.Enabled = False
            '.cmdSaveDef.Enabled = False
        End If
    End With
    
    bScanMode = False
    SectionOutOfLimit "", bErase:=True
    
    Dim sEDS_Time   As String
    Dim OSData      As String
    
    If bDebugMode Or bDebugToFile Then
    
        If ObjPtr(OSver) <> 0 Then
                OSData = OSver.Bitness & " " & OSver.OSName & IIf(OSver.Edition <> "", " (" & OSver.Edition & ")", "") & ", " & _
                    OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & ", " & _
                    "Service Pack: " & Replace(OSver.SPVer, ",", ".") & IIf(OSver.IsSafeBoot, " (Safe Boot)", "")
        End If
    
        sEDS_Time = vbCrLf & vbCrLf & "Logging is finished." & vbCrLf & vbCrLf & AppVerPlusName & vbCrLf & vbCrLf & OSData & vbCrLf & vbCrLf & _
                "Time spent: " & ((GetTickCount() - Perf.StartTime) \ 100) / 10 & " sec." & vbCrLf & vbCrLf & _
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
        If g_hDebugLog <> 0 Then
            'Append Header to the end and close debug log file
            Dim b() As Byte
            b = sEDS_Time & vbCrLf & vbCrLf
            PutW g_hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
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
    
    If g_bCheckSum Then
        frmMain.lblMD5.Visible = True
        frmMain.shpMD5Background.Visible = True
        frmMain.shpMD5Progress.Visible = True
    End If
    
    'ProgressBar label settings
    frmMain.lblStatus.Visible = True
    frmMain.lblStatus.Caption = ""
    frmMain.lblStatus.ForeColor = &HFFFF&   'Yellow
    frmMain.lblStatus.ZOrder 0 'on top
    frmMain.lblStatus.Left = 400
    
    'Logo -> off
    frmMain.pictLogo.Visible = False
    
    'results label -> off
    frmMain.lblInfo(0).Visible = False
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

Public Sub CloseProgressbar(Optional bScanCompeleted As Boolean)
    frmMain.shpBackground.Visible = False
    frmMain.shpProgress.Visible = False
    frmMain.shpMD5Background.Visible = False
    frmMain.shpMD5Progress.Visible = False
    frmMain.lblStatus.Visible = False
    frmMain.lblMD5.Visible = False
    If bScanCompeleted Then
        If frmMain.lstResults.Visible Then
            frmMain.lblInfo(1).Visible = True
        End If
        If Not TaskBar Is Nothing Then TaskBar.SetProgressState g_HwndMain, TBPF_NOPROGRESS
    End If
End Sub

Public Sub ResumeProgressbar()
    frmMain.shpBackground.Visible = True
    frmMain.shpProgress.Visible = True
    frmMain.lblStatus.Visible = True
    If g_bCheckSum Then
        frmMain.shpMD5Background.Visible = True
        frmMain.shpMD5Progress.Visible = True
        frmMain.lblMD5.Visible = True
    End If
End Sub

Public Sub UpdateProgressBar(Section As String, Optional sAppendText As String)
    On Error GoTo ErrorHandler:
    
    If g_bNoGUI Then Exit Sub
    
    AppendErrorLogCustom "Progressbar - " & Section & " " & sAppendText
    
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
            Case "O7-Cert": .lblStatus.Caption = Translate(264) & "..."
            Case "O7-Trouble": .lblStatus.Caption = Translate(265) & "..."
            Case "O7-ACL": .lblStatus.Caption = Translate(267) & "..."
            Case "O7-IPSec": .lblStatus.Caption = Translate(266) & "..."
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
            Case "O23-D": .lblStatus.Caption = Translate(263) & "..."
            Case "O24": .lblStatus.Caption = Translate(257) & "..."
            Case "O25": .lblStatus.Caption = Translate(258) & "..."
            Case "O26": .lblStatus.Caption = Translate(261) & "..."
            
            Case "ProcList": .lblStatus.Caption = Translate(260) & "..."
            Case "ModuleList": .lblStatus.Caption = Translate(268) & "..."
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
    
    Dim vRule As Variant, iMode&, bIsNSBSD As Boolean, result As SCAN_RESULT
    Dim sHit$, sKey$, sParam$, sData$, sDefDataStrings$, Wow6432Redir As Boolean, UseWow
    Dim bProxyEnabled As Boolean, hHive As ENUM_REG_HIVE
    
    'Registry rule syntax:
    '[regkey],[regvalue],[infected data],[default data]
    '* [regkey]           = "" -> abort - no way man!
    ' * [regvalue]        = "" -> delete entire key
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
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey sKey
    
    Do While HE.MoveNext
    
        Wow6432Redir = HE.Redirected
        sKey = HE.Key
        hHive = HE.Hive
    
        Select Case iMode
        
        Case 0 'check for incorrect value
            If Reg.ValueExists(hHive, sKey, sParam, Wow6432Redir) Then
            
                sData = Reg.GetString(hHive, sKey, sParam, Wow6432Redir)
              
                If Len(sData) <> 0 Then
              
                    sData = UnQuote(EnvironW(sData))
            
                    If Not inArraySerialized(sData, sDefDataStrings, "|", , , 1) Or (Not bHideMicrosoft) Then
                        bIsNSBSD = False
                        If bHideMicrosoft And Not bIgnoreAllWhitelists Then bIsNSBSD = StrBeginWithArray(sData, aSafeRegDomains)
                        If Not bIsNSBSD Then
                            If InStr(1, sData, "%2e", 1) > 0 Then sData = UnEscape(sData)
                    
                            sHit = IIf(bIsWin32, "R0 - ", IIf(Wow6432Redir, "R0-32 - ", "R0 - ")) & _
                                HE.KeyAndHive & ": " & IIf(sParam = "", "(default)", "[" & sParam & "]") & _
                                " = " & IIf(sData <> "", sData, "(empty)") 'doSafeURLPrefix
                    
                            If Not IsOnIgnoreList(sHit) Then
                                With result
                                    .Section = "R0"
                                    .HitLineW = sHit
                                    AddRegToFix .Reg, RESTORE_VALUE, hHive, sKey, sParam, SplitSafe(sDefDataStrings, "|")(0), CLng(Wow6432Redir)
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults result
                            End If
                        End If
                    End If
                End If
            End If
            
        Case 1  'check for present value
            
            If Reg.ValueExists(hHive, sKey, sParam, Wow6432Redir) Then
            
                sData = Reg.GetString(hHive, sKey, sParam, Wow6432Redir)
              
                If Len(sData) <> 0 Then
            
                    'check if domain is on safe list
                    bIsNSBSD = False
                    If bHideMicrosoft And Not bIgnoreAllWhitelists Then bIsNSBSD = StrBeginWithArray(sData, aSafeRegDomains)
                    'make hit
                    If Not bIsNSBSD Then
                        If InStr(1, sData, "%2e", 1) > 0 Then sData = UnEscape(sData)

                        If sParam = "ProxyServer" Then
                            bProxyEnabled = (Reg.GetDword(hHive, sKey, "ProxyEnable", Wow6432Redir) = 1)
                            
                            sHit = IIf(bIsWin32, "R1 - ", IIf(Wow6432Redir, "R1-32 - ", "R1 - ")) & _
                                HE.KeyAndHive & ": " & IIf(sParam = "", "(default)", "[" & sParam & "]") & " = " & _
                                IIf(sData <> "", sData, "(empty)") & IIf(bProxyEnabled, " (enabled)", " (disabled)")
                        Else
                            sHit = IIf(bIsWin32, "R1 - ", IIf(Wow6432Redir, "R1-32 - ", "R1 - ")) & _
                                HE.KeyAndHive & ": " & IIf(sParam = "", "(default)", "[" & sParam & "]") & " = " & _
                                IIf(sData <> "", sData, "(empty)")   'doSafeURLPrefix
                        End If
                    
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "R1"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REMOVE_VALUE, hHive, sKey, sParam, , CLng(Wow6432Redir)
                                If sParam = "ProxyServer" Then
                                    AddRegToFix .Reg, RESTORE_VALUE, hHive, sKey, "ProxyEnable", 0, CLng(Wow6432Redir), REG_RESTORE_DWORD
                                End If
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                End If
            End If
            
        Case 2 'check if regkey is present
            If Reg.KeyExists(hHive, sKey, Wow6432Redir) Then
            
                sHit = IIf(bIsWin32, "R2 - ", IIf(Wow6432Redir, "R2-32 - ", "R2 - ")) & HE.KeyAndHive
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "R2"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, hHive, sKey, , , CLng(Wow6432Redir)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
            End If
        End Select
    Loop
    
    'Set HE = Nothing
    
    AppendErrorLogCustom "ProcessRuleReg - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_ProcessRuleReg", "sRule=", sRule
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixRegItem(sItem$, result As SCAN_RESULT)
    'R0 - HKCU\Software\..\Main,Window Title
    'R1 - HKCU\Software\..\Main,Window Title=MSIE 5.01
    'R2 - HKCU\Software\..\Main
    FixRegistryHandler result
End Sub


'CheckR3item
Public Sub CheckR3Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckR3Item - Begin"

    Dim sURLHook$, hKey&, i&, sName$, sHit$, sCLSID$, sFile$, result As SCAN_RESULT, lret&
    Dim bHookMising As Boolean, sDefHookDll$, sDefHookCLSID$, sHookDll_1$, sHookDll_2$
    
    sURLHook = "Software\Microsoft\Internet Explorer\URLSearchHooks"
    
    sDefHookCLSID = "{CFBFAE00-17A6-11D0-99CB-00C04FD64497}"

    sHookDll_1 = sWinSysDir & "\ieframe.dll"
    sHookDll_2 = sWinSysDir & "\shdocvw.dll"

    If OSver.MajorMinor >= 5.2 Then 'XP x64 +
        sDefHookDll = sHookDll_1
    Else
        sDefHookDll = sHookDll_2
    End If
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_HKCU Or HE_HIVE_HKU, HE_SID_USER Or HE_SID_NO_VIRTUAL, HE_REDIR_NO_WOW
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
                With result
                    .Section = "R3"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sDefHookCLSID, vbNullString, , REG_RESTORE_SZ
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID, "", "Microsoft Url Search Hook"
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID & "\InProcServer32", "", sDefHookDll
                    AddRegToFix .Reg, RESTORE_VALUE, HKCR, "CLSID\" & sDefHookCLSID & "\InProcServer32", "ThreadingModel", "Apartment"
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Loop
    
    HE.Init HE_HIVE_ALL
    HE.AddKey sURLHook
    
    Do While HE.MoveNext
        
        lret = RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0&, KEY_QUERY_VALUE Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey)
        
        If lret = 0 Then
        
          sCLSID = String$(MAX_VALUENAME, 0&)
          i = 0
          Do While 0 = RegEnumValueW(hKey, i, StrPtr(sCLSID), Len(sCLSID), 0&, ByVal 0&, 0&, ByVal 0&)
            
            sCLSID = TrimNull(sCLSID)
            
            GetFileByCLSID sCLSID, sFile, , HE.Redirected, HE.SharedKey
            
            If Not (sCLSID = sDefHookCLSID And _
                (StrComp(sFile, sHookDll_1, 1) = 0 Or StrComp(sFile, sHookDll_2, 1) = 0)) Or (Not bHideMicrosoft) Then
                
                GetTitleByCLSID sCLSID, sName, HE.Redirected, HE.SharedKey
                
                sHit = IIf(bIsWin32, "R3 - ", IIf(HE.Redirected, "R3-32 - ", "R3 - ")) & HE.HiveNameAndSID & "\..\URLSearchHooks: " & _
                    sName & " - " & sCLSID & " - " & sFile
                    
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "R3"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
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

Public Sub FixR3Item(sItem$, result As SCAN_RESULT)
    'R3 - Shitty search hook - {00000000} - c:\windows\bho.dll"
    'R3 - Default URLSearchHook is missing
    
    FixRegistryHandler result
End Sub

Public Sub CheckR4Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckR4Item - Begin"
    
    'http://ijustdoit.eu/changing-default-search-provider-in-internet-explorer-11-using-group-policies/
    
    'SearchScope
    'R4 - SearchScopes:
    
    Dim result As SCAN_RESULT, sHit$, j&, k&, sURL$, sProvider$, aScopes() As String, sBuf$, sDefScope$, bDefault As Boolean
    Dim Param As Variant, aData() As String, sHive$, sParam$, sDefData$
    Dim HE As clsHiveEnum, HEFix As clsHiveEnum
    Set HE = New clsHiveEnum
    Set HEFix = New clsHiveEnum
    
    'Enum custom scopes
    '
    'HKCU\Software\Policies\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Policies\Microsoft\Internet Explorer\SearchScopes
    'HKLM\Software\Microsoft\Internet Explorer\SearchScopes
    'HKCU\Software\Microsoft\Internet Explorer\SearchScopes
    
    HE.Init HE_HIVE_ALL, (HE_SID_ALL And Not HE_SID_SERVICE) Or HE_SID_NO_VIRTUAL, HE_REDIR_NO_WOW
    
    HE.AddKey "Software\Microsoft\Internet Explorer\SearchScopes"
    HE.AddKey "Software\Policies\Microsoft\Internet Explorer\SearchScopes"
    
    HE.Clone HEFix
    
    Dim sLastURL As String
    Dim sParams As String
    
    Do While HE.MoveNext
        
        For j = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aScopes())
            
            sProvider = Reg.GetString(HE.Hive, HE.Key & "\" & aScopes(j), "DisplayName")
            If sProvider = "" Then sProvider = "(no name)"
            
            If Left$(sProvider, 1) = "@" Then
                sBuf = GetStringFromBinary(, , sProvider)
                If 0 <> Len(sBuf) Then sProvider = sBuf
            End If
            
            sParams = ""
            sLastURL = ""
            
            For Each Param In Array("URL", "SuggestionsURL_JSON", "SuggestionsURL", "SuggestionsURLFallback", "TopResultURL", "TopResultURLFallback")
            
              sURL = Reg.GetString(HE.Hive, HE.Key & "\" & aScopes(j), CStr(Param))
              
              If Len(sURL) <> 0 Or Reg.ValueExists(HE.Hive, HE.Key & "\" & aScopes(j), CStr(Param)) Then
                
                If Not IsBingScopeKeyPara("URL", sURL) Then
                  
                    With result
                        .Section = "R4"
                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aScopes(j)
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
                            sHit = "R4 - SearchScopes: " & HE.KeyAndHive & "\" & aScopes(j) & ": [" & sParams & "] = " & sLastURL & " - " & sProvider
                            
                            If Not IsOnIgnoreList(sHit) Then
                                result.HitLineW = sHit
                                GoSub CheckDefaultScope
                                AddToScanResults result, , True
                            End If
                            
                            sLastURL = sURL
                            sParams = CStr(Param)
                        End If
                    End If
                End If
              End If
            Next
            
            If sParams <> "" And Len(sLastURL) <> 0 Then
                sHit = "R4 - SearchScopes: " & HE.KeyAndHive & "\" & aScopes(j) & ": [" & sParams & "] = " & sLastURL & " - " & sProvider
                
                If Not IsOnIgnoreList(sHit) Then
                    result.HitLineW = sHit
                    GoSub CheckDefaultScope
                    AddToScanResults result
                End If
            End If
            
        Next
    Loop
    
    AppendErrorLogCustom "CheckR4Item - End"
    
    'Set HE = Nothing
    Set HEFix = Nothing
    
    Exit Sub
    
CheckDefaultScope:
    
    HEFix.Repeat
    Do While HEFix.MoveNext
        sDefScope = Reg.GetString(HEFix.Hive, HEFix.Key, "DefaultScope")
        If sDefScope <> "" Then
            If StrComp(sDefScope, aScopes(j), 1) = 0 Then
                If InStr(1, HEFix.Key, "Policies", 1) <> 0 Then
                    'remove policies
                    AddRegToFix result.Reg, REMOVE_VALUE, HEFix.Hive, HEFix.Key, "DefaultScope"
                Else
                    'reset default scope to bing
                    
                    For k = 1 To cReg4vals.Count
                        aData = Split(cReg4vals.Item(k), ",", 3)
                        sHive = aData(0)
                        sParam = aData(1)
                        sDefData = SplitSafe(aData(2), "|")(0)
        
                        If (HEFix.HiveName = "HKLM" And sHive = "HKLM") Or _
                            (HEFix.HiveName <> "HKLM" And sHive = "HKCU") Then
                
                            AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", sParam, sDefData, , REG_RESTORE_SZ
                        End If
                    Next
            
                    AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key, "DefaultScope", "{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", , REG_RESTORE_SZ
                    AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "DisplayName", "Bing", , REG_RESTORE_SZ
                    If HEFix.HiveName = "HKLM" Then
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "", "Bing", , REG_RESTORE_SZ
                    Else
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "FaviconURL", "https://www.bing.com/favicon.ico", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "FaviconURLFallback", "https://www.bing.com/favicon.ico", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTLogoPath", AppDataLocalLow & "\Microsoft\Internet Explorer\Services\", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTLogoURL", "https://go.microsoft.com/fwlink/?LinkID=403856&language={language}&scale={scalelevel}&contrast={contrast}", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTSuggestionsURL", "https://api.bing.com/qsml.aspx?query={searchTerms}&market={language}&maxwidth={ie:maxWidth}&rowheight={ie:rowHeight}&sectionHeight={ie:sectionHeight}&FORM=IENTSS", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTTopResultURL", "https://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTTR", , REG_RESTORE_SZ
                        AddRegToFix result.Reg, RESTORE_VALUE, HEFix.Hive, HEFix.Key & "\{0633EE93-D776-472f-A0FF-E1416B8B2E3A}", "NTURL", "https://www.bing.com/search?q={searchTerms}&src=IE-SearchBox&FORM=IENTSR", , REG_RESTORE_SZ
                    End If
                End If
            End If
        End If
    Loop
    
    Return
ErrorHandler:
    ErrorMsg Err, "CheckR4Item"
    If inIDE Then Stop: Resume Next
End Sub

Private Function IsBingScopeKeyPara(sRegParam As String, sURL As String) As Boolean
    If sURL = "" Then Exit Function
    
    If Not bHideMicrosoft Then Exit Function
    
    'Is valid domain
    Dim pos As Long, sPrefix As String
    pos = InStr(sURL, "?")
    If pos = 0 Then Exit Function
    sPrefix = Left$(sURL, pos - 1)
    Select Case sPrefix
    Case "http://search.microsoft.com/results.aspx"
    Case "http://www.bing.com/search"
    Case "http://www.bing.com/as/api/qsml"
    Case "http://api.bing.com/qsml.aspx"
    Case "http://search.live.com/results.aspx"
    Case "http://api.search.live.com/qsml.aspx"
    Case Else
        Exit Function
    End Select

    Dim aKey() As String, aVal() As String, i As Long
    Dim bSearchTermPresent As Boolean
    
    IsBingScopeKeyPara = True
    
    If StrEndWith(sURL, ";") Then sURL = Left$(sURL, Len(sURL) - 1)
    
    Call ParseKeysURL(sURL, aKey, aVal)
    
    Select Case UCase$(sRegParam)
    
        Case "URL", UCase$("SuggestionsURL"), UCase$("SuggestionsURLFallback"), UCase$("TopResultURL"), UCase$("TopResultURLFallback")
            If AryItems(aKey) Then
                For i = 0 To UBound(aKey)
                    Select Case LCase(aKey(i))
                    Case "q", "query"
                    '{searchTerms}
                        If StrComp(aVal(i), "{searchTerms}", 1) = 0 Then bSearchTermPresent = True
                    
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
                    
                    Case "mkt"
                    Case "setlang"
                    Case "ptag"
                    Case "conlogo"
                    
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

Public Sub FixR4Item(sItem$, result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    FixRegistryHandler result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixR4Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Private Sub CheckFileItems(ByVal sRule$)
    On Error GoTo ErrorHandler:
    
    Dim vRule As Variant, iMode&, sHit$, result As SCAN_RESULT
    Dim sFile$, sSection$, sParam$, sData$, sLegitData$
    Dim sTmp$
    
    AppendErrorLogCustom "CheckFileItems - Begin", "Rule: " & sRule
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
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
            sData = Trim$(RTrimNull(sData))
            
            If bIsWinNT And Len(sData) <> 0 Then
            
                If Not inArraySerialized(sData, sLegitData, "|", , , vbTextCompare) Or (Not bHideMicrosoft) Then
                
                    sHit = "F0 - " & sFile & ": " & "[" & sSection & "]" & " " & sParam & " = " & sData
                    If Not IsOnIgnoreList(sHit) Then
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
                        With result
                            .Section = "F0"
                            .HitLineW = sHit
                            'system.ini
                            AddIniToFix .Reg, RESTORE_VALUE_INI, sFile, "boot", "shell", SplitSafe(sLegitData, "|")(0)  '"explorer.exe"
                            .CureType = INI_BASED
                        End With
                        AddToScanResults result
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
            sData = Trim$(RTrimNull(sData))
            
            If Len(sData) <> 0 Then
                sHit = "F1 - " & sFile & ": " & "[" & sSection & "]" & " " & sParam & " = " & sData
                If Not IsOnIgnoreList(sHit) Then
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
                    With result
                        .Section = "F1"
                        .HitLineW = sHit
                        'win.ini
                        AddIniToFix .Reg, RESTORE_VALUE_INI, sFile, "windows", sParam, "" 'param = 'load' or 'run'
                        .CureType = INI_BASED
                    End With
                    AddToScanResults result
                End If
            End If
            
        Case 2
            'F2 = check if value is infected, in the Registry
            'so far F2 is only reg:Shell and reg:UserInit
            
            HE.Init HE_HIVE_ALL
            HE.AddKey "Software\Microsoft\Windows NT\CurrentVersion\WinLogon"
            
            Do While HE.MoveNext
                
                sData = Reg.GetString(HE.Hive, HE.Key, sParam, HE.Redirected)
                sTmp = sData
                If Right$(sData, 1) = "," Then sTmp = Left$(sTmp, Len(sTmp) - 1)
                
                'Note: HKCU + empty values are allowed
                If (Not inArraySerialized(sTmp, sLegitData, "|", , , vbTextCompare) Or (Not bHideMicrosoft)) And _
                  Not ((HE.Hive = HKCU Or HE.Hive = HKU) And sData = "") Then
            
                    'exclude no WOW64 value on Win10 for UserInit
                    If Not (HE.Redirected And OSver.MajorMinor >= 10 And sParam = "UserInit" And sData = "") Then
                    If Not (HE.Redirected And sParam = "UserInit" And StrComp(sData, BuildPath(sWinSysDirWow64, "userinit.exe")) = 0) Then
                
                        sHit = IIf(bIsWin32, "F2 - ", IIf(HE.Redirected, "F2-32 - ", "F2 - ")) & HE.HiveNameAndSID & "\..\WinLogon: " & _
                            "[" & sParam & "] = " & sData
                        If Not IsOnIgnoreList(sHit) Then
                            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
                            With result
                                .Section = "F2"
                                .HitLineW = sHit
                                AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sParam, SplitSafe(sLegitData, "|")(0), HE.Redirected
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                    End If
                End If
            Loop
            
        Case 3
            'F3 = check if value is present, in the Registry
            'this is not really smart when more INIFile items get
            'added, but so far F3 is only reg:load and reg:run
        
            HE.Init HE_HIVE_ALL
            HE.AddKey "Software\Microsoft\Windows NT\CurrentVersion\Windows"
            
            Do While HE.MoveNext
            
                sData = Reg.GetString(HE.Hive, HE.Key, sParam, HE.Redirected)
                If 0 <> Len(sData) Then
                    sHit = IIf(bIsWin32, "F3 - ", IIf(HE.Redirected, "F3-32 - ", "F3 - ")) & HE.HiveNameAndSID & "\..\Windows: " & _
                        "[" & sParam & "] = " & sData
                    If Not IsOnIgnoreList(sHit) Then
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sData)
                        With result
                            .Section = "F3"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
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

Public Sub FixFileItem(sItem$, result As SCAN_RESULT)
    'F0 - system.ini: Shell=c:\win98\explorer.exe openme.exe
    'F1 - win.ini: load=hpfsch
    'F2, F3 - registry

    'coding is easy if you cheat :)
    '(c) Dragokas: Cheaters will be punished ^_^
    
    FixRegistryHandler result
End Sub

Public Sub CheckO1Item_DNSApi()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item_DNSApi - Begin"
    
    If OSver.MajorMinor <= 5 Then Exit Sub 'XP+ only
    
    Const MaxSize As Long = 5242880 ' 5 MB.
    
    Dim vFile As Variant, ff As Long, Size As Currency, p As Long, buf() As Byte, sHit As String, result As SCAN_RESULT
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
                
                    p = InArrSign_NoCase(0, buf, bufExample, bufExample_2)
                    
                    If p = -1 Then                      '//TODO: add isMicrosoftFile() ?
                        ' if signature not found
                        sHit = "O1 - DNSApi: File is patched - " & vFile
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O1"
                                .HitLineW = sHit
                                AddFileToFix .File, RESTORE_FILE_SFC, CStr(vFile)
                                .CureType = FILE_BASED
                            End With
                            AddToScanResults result
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

Public Function InArrString(pos As Long, ArrSrc() As Byte, StrExample As Byte, CompareMethod As VbCompareMethod) As Long
    If CompareMethod = vbBinaryCompare Then
        Dim aExample() As Byte
        aExample = StrConv(StrExample, vbFromUnicode)
        InArrString = InArrSign(pos, ArrSrc, aExample)
    Else
        Dim aLCase() As Byte
        Dim aUCase() As Byte
        aLCase = StrConv(LCase$(StrExample), vbFromUnicode)
        aUCase = StrConv(UCase$(StrExample), vbFromUnicode)
        InArrString = InArrSign_NoCase(pos, ArrSrc, aLCase, aUCase)
    End If
End Function

Private Function InArrSign(pos As Long, ArrSrc() As Byte, ArrEx() As Byte) As Long
    Dim i As Long, j As Long, p As Long, Found As Boolean
    InArrSign = -1
    For i = pos To UBound(ArrSrc) - UBound(ArrEx)
        p = i
        Found = True
        For j = 0 To UBound(ArrEx)
            If ArrSrc(p) <> ArrEx(j) Then Found = False: Exit For
            p = p + 1
        Next
        If Found Then InArrSign = p - UBound(ArrEx) - 1: Exit For
    Next
End Function

Private Function InArrSign_NoCase(pos As Long, ArrSrc() As Byte, ArrEx() As Byte, ArrEx_2() As Byte) As Long
    'ArrEx - all lcase
    'ArrEx_2 - all Ucase
    Dim i As Long, j As Long, p As Long, Found As Boolean
    InArrSign_NoCase = -1
    For i = pos To UBound(ArrSrc) - UBound(ArrEx)
        p = i
        Found = True
        For j = 0 To UBound(ArrEx)
            If ArrSrc(p) <> ArrEx(j) And _
                ArrSrc(p) <> ArrEx_2(j) Then Found = False: Exit For
            p = p + 1
        Next
        If Found Then InArrSign_NoCase = p - UBound(ArrEx) - 1: Exit For
    Next
End Function

Public Sub CheckO1Item_ICS()
    ' hosts.ics
    'https://support.microsoft.com/ru-ru/kb/309642
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item_ICS - Begin"
    
    Dim sHostsFileICS$, sHit$, sHostsFileICS_Default$
    Dim sLines$, sLine As Variant, NonDefaultPath As Boolean, cFileSize As Currency, hFile As Long
    Dim result As SCAN_RESULT
    
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
    
    If OpenW(sHostsFileICS, FOR_READ, hFile, g_FileBackupFlag) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
        ToggleWow64FSRedirection True
    Else
    
        sHit = "O1 - Unable to read Hosts.ICS file"
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFileICS
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
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
                    With result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, sHostsFileICS
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
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
    
    If OpenW(sHostsFileICS_Default, FOR_READ, hFile, g_FileBackupFlag) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
    Else
        sHit = "O1 - Unable to read Hosts.ICS default file"
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFileICS_Default
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
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
                    With result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, sHostsFileICS_Default
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
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


Public Sub CheckO1Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO1Item - Begin"
    
    Dim sHit$, i&, ff%, HostsDefaultFile$, NonDefaultPath As Boolean
    Dim sLine As Variant, sLines$, cFileSize@
    Dim aHits() As String, j As Long, hFile As Long
    Dim HostsDefaultPath As String
    ReDim aHits(0)
    Dim result As SCAN_RESULT
    
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
            With result
                .Section = "O1"
                .HitLineW = sHit
                HostsDefaultPath = EnvironUnexpand(GetParentDir(HostsDefaultFile))
                AddRegToFix .Reg, RESTORE_VALUE, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Tcpip\Parameters", "DatabasePath", _
                  HostsDefaultPath, , REG_RESTORE_EXPAND_SZ
                .CureType = REGISTRY_BASED Or CUSTOM_BASED
            End With
            AddToScanResults result
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
            sHit = "O1 - Hosts: is empty"

            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O1"
                    .HitLineW = sHit
                    AddFileToFix .File, BACKUP_FILE, sHostsFile
                    .CureType = CUSTOM_BASED
                End With
                AddToScanResults result
            End If
            
            ToggleWow64FSRedirection True
            Exit Sub
        End If
    End If
    
    Dbg "5"
    
    If OpenW(sHostsFile, FOR_READ, hFile, g_FileBackupFlag) Then
        sLines = String$(cFileSize, vbNullChar)
        GetW hFile, 1, sLines
        CloseW hFile
        ToggleWow64FSRedirection True
    Else
    
        sHit = "O1 - Hosts: Unable to read Hosts file"
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFile
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
        
        ToggleWow64FSRedirection True
        If NonDefaultPath Then
            GoTo CheckHostsDefault:
        Else
            Exit Sub
        End If
    End If
    
    sLines = Replace$(sLines, vbCrLf, vbLf)
    
    If Len(Replace$(sLines, vbNullChar, "")) = 0 Then
    
        sHit = "O1 - Hosts: is damaged (contains NUL characters only)"
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFile
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
        
        ToggleWow64FSRedirection True
        If NonDefaultPath Then
            GoTo CheckHostsDefault:
        Else
            Exit Sub
        End If
    End If
    
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
                If (Left$(sLine, 1) <> "#" Or bIgnoreAllWhitelists) And _
                  ((StrComp(sLine, "127.0.0.1       localhost", 1) <> 0 And _
                  StrComp(sLine, "::1             localhost", 1) <> 0 And _
                  StrComp(sLine, "127.0.0.1 localhost", 1) <> 0) Or Not bHideMicrosoft) Then
                  
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
                    With result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, sHostsFile
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
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
            With result
                .Section = "O1"
                .HitLineW = sHit
                AddFileToFix .File, BACKUP_FILE, sHostsFile
                .CureType = CUSTOM_BASED
            End With
            'limit for first and last 20 entries only to view on results window
            AddToScanResults result, IIf((j < 20) Or (j > i - 1 - 20), False, True)
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

                If OpenW(HostsDefaultFile, FOR_READ, hFile, g_FileBackupFlag) Then
                    sLines = String$(cFileSize, vbNullChar)
                    GetW hFile, 1, sLines
                    CloseW hFile
                Else
                    sHit = "O1 - Hosts default: Unable to read Default Hosts file"

                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O1"
                            .HitLineW = sHit
                            AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                            .CureType = CUSTOM_BASED
                        End With
                        AddToScanResults result
                    End If
                    
                    Exit Sub
                End If
                
                Dbg "10"
                
                sLines = Replace$(sLines, vbCrLf, vbLf)
                
                If Len(Replace$(sLines, vbNullChar, "")) = 0 Then
    
                    sHit = "O1 - Hosts default: is damaged (contains NUL characters only)"
        
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O1"
                            .HitLineW = sHit
                            AddFileToFix .File, BACKUP_FILE, sHostsFile
                            .CureType = CUSTOM_BASED
                        End With
                        AddToScanResults result
                    End If
                    
                    Exit Sub
                End If

                For Each sLine In Split(sLines, vbLf)
                
                    sLine = Replace$(sLine, vbTab, " ")
                    sLine = Replace$(sLine, vbCr, "")
                    sLine = Trim$(sLine)
                    
                    If sLine <> vbNullString Then
                    
                        If (Left$(sLine, 1) <> "#" Or bIgnoreAllWhitelists) And _
                          ((StrComp(sLine, "127.0.0.1       localhost", 1) <> 0 And _
                          StrComp(sLine, "::1             localhost", 1) <> 0) Or Not bHideMicrosoft) Then    '::1 - default for Vista
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
                    With result
                        .Section = "O1"
                        .HitLineW = sHit
                        AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                        .CureType = CUSTOM_BASED
                    End With
                    AddToScanResults result
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
                With result
                    .Section = "O1"
                    .HitLineW = sHit
                    AddFileToFix .File, BACKUP_FILE, HostsDefaultFile
                    .CureType = CUSTOM_BASED
                End With
                'limit for first and last 20 entries only to view on results window
                AddToScanResults result, IIf((j < 20) Or (j > i - 1 - 20), False, True)
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

Public Sub FixO1Item(sItem$, result As SCAN_RESULT)
    'O1 - Hijack of auto.search.msn.com etc with Hosts file
    On Error GoTo ErrorHandler:
    Dim sLine As Variant, sHijacker$, i&, iAttr&, ff1%, ff2%, HostsDefaultPath$, sLines$, HostsDefaultFile$, cFileSize@, sHosts$
    Dim sHostsTemp$, bResetHosts As Boolean, aLines() As String, isICS As Boolean, SFC As String
    
    If InStr(1, sItem, "O1 - DNSApi:", 1) <> 0 Then
        FixFileHandler result
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
    
    If StrComp(sItem, "O1 - Hosts: is empty", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts: Unable to read Hosts file", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts default: Unable to read Default Hosts file", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts: Reset contents to default", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts default: Reset contents to default", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts: is damaged (contains NUL characters only)", 1) = 0 Or _
      StrComp(sItem, "O1 - Hosts default: is damaged (contains NUL characters only)", 1) = 0 Then
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
    If Not CheckAccessWrite(sHostsTemp) Then
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
    
    BackupFile result, sHosts
    
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
        If Not bAutoLogSilent Then
            MsgBoxW Translate(303), vbExclamation
        End If
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
    
    Dim hKey&, i&, sName$, sCLSID$, lpcName&, sFile$, sHit$, BHO_key$, result As SCAN_RESULT
    Dim sBuf$, sProgId$, sProgId_CLSID$, bSafe As Boolean
    
    Dim HEFixKey As clsHiveEnum
    Dim HEFixValue As clsHiveEnum
    
    Set HEFixKey = New clsHiveEnum
    Set HEFixValue = New clsHiveEnum
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HEFixKey.Init HE_HIVE_ALL
    HEFixValue.Init HE_HIVE_ALL
    
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
                    
                    GetFileByCLSID sCLSID, sFile, , HE.Redirected, HE.SharedKey
                    
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
                    
                    sFile = FormatFileMissing(sFile)
                    
                    bSafe = False
                    If bHideMicrosoft Then
                        If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                            If IsMicrosoftFile(sFile) Then bSafe = True
                        End If
                    End If
                    
                    If Not bSafe Then
                        'get bho name from CLSID
                        If sName = "" Then GetTitleByCLSID sCLSID, sName, HE.Redirected, HE.SharedKey
                    
                        sHit = IIf(bIsWin32, "O2", IIf(HE.Redirected, "O2-32", "O2")) & _
                            " - " & HE.HiveNameAndSID & "\..\BHO: " & sName & " - " & sCLSID & " - " & sFile
                    
                        If Not IsOnIgnoreList(sHit) Then
                        
                            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                            
                            With result
                                .Section = "O2"
                                .HitLineW = sHit
                                
                                AddRegToFix .Reg, REMOVE_KEY, 0, BHO_key, , , IIf(HE.SharedKey, REG_REDIRECTION_BOTH, HE.Redirected)
                    
                                If 0 <> Len(sProgId) Then
                                    AddRegToFix .Reg, REMOVE_KEY, HKCR, sProgId, , , IIf(HE.SharedKey, REG_REDIRECTION_BOTH, HE.Redirected)
                                End If
                                
                                HEFixKey.Repeat
                                Do While HEFixKey.MoveNext
                                    AddRegToFix .Reg, REMOVE_KEY, HEFixKey.Hive, Replace$(HEFixKey.Key, "{CLSID}", sCLSID), , , HEFixKey.Redirected
                                Loop
                                
                                HEFixValue.Repeat
                                Do While HEFixValue.MoveNext
                                    AddRegToFix .Reg, REMOVE_VALUE, HEFixValue.Hive, HEFixValue.Key, sCLSID, , HEFixValue.Redirected
                                Loop
                                
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile 'Or UNREG_DLL
                                
                                .CureType = REGISTRY_BASED Or FILE_BASED
                            End With
                        
                            AddToScanResults result
                        End If
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

Public Sub FixO2Item(sItem$, result As SCAN_RESULT)
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
    FixFileHandler result
    FixRegistryHandler result
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
    
    Dim hKey&, i&, sCLSID$, sName$, result As SCAN_RESULT
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
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HEFixKey.Init HE_HIVE_ALL
    HEFixValue.Init HE_HIVE_ALL
    
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
                GetFileByCLSID sCLSID, sFile, , HE.Redirected, HE.SharedKey
    
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
                
                sFile = FormatFileMissing(sFile)
                
                bSafe = False
                
                If bHideMicrosoft Then
                    If OSver.MajorMinor = 5 Then 'Win2k
                        If WhiteListed(sFile, sWinDir & "\system32\msdxm.ocx") Then bSafe = True
                    End If
                End If
                
                'If 0 <> Len(sName) And InStr(sCLSID, "{") > 0 And Not bSafe Then
                If InStr(sCLSID, "{") <> 0 And Not bSafe Then
    
    '          If Not SearchwwwTrick Or _
    '            (SearchwwwTrick And (sCLSID <> "BrandBitmap" And sCLSID <> "SmBrandBitmap")) Then
                    
                    GetTitleByCLSID sCLSID, sName, HE.Redirected, HE.SharedKey
    
                    sHit = IIf(bIsWin32, "O3", IIf(HE.Redirected, "O3-32", "O3")) & _
                        " - " & HE.HiveNameAndSID & "\..\" & aDescr(HE.KeyIndex) & ": " & sName & " - " & sCLSID & " - " & sFile
                    
                    If Not IsOnIgnoreList(sHit) Then
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                        With result
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
                                AddRegToFix .Reg, REMOVE_KEY, HEFixKey.Hive, Replace$(HEFixKey.Key, "{CLSID}", sCLSID), , , HEFixKey.Redirected
                            Loop
                            
                            HEFixValue.Repeat
                            Do While HEFixValue.MoveNext
                                AddRegToFix .Reg, REMOVE_VALUE, HEFixValue.Hive, HEFixValue.Key, sCLSID, , HEFixValue.Redirected
                            Loop
                            
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile 'Or UNREG_DLL
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
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

Public Sub FixO3Item(sItem$, result As SCAN_RESULT)
    'O3 - Enumeration of existing MSIE toolbars

    FixFileHandler result
    FixRegistryHandler result
End Sub


'returns array of SID strings, except of current user
Sub GetUserNamesAndSids(aSID() As String, aUser() As String)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetUserNamesAndSids - Begin"
    
    'get all users' SID and map it to the corresponding username
    'not all users visible in User Accounts screen have a SID in HKU hive though,
    'they get it when logged

    Dim CurUserName$, i&, k&, sUsername$, aTmpSID() As String, aTmpUser() As String

    CurUserName = OSver.UserName
    
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
    k = 0
    ReDim aSID(UBound(aTmpSID))
    ReDim aUser(UBound(aTmpSID))
    
    For i = 0 To UBound(aTmpSID)
        If 0 <> Len(aTmpSID(i)) Then
            aSID(k) = aTmpSID(i)
            aUser(k) = aTmpUser(i)
            k = k + 1
        End If
    Next
    If k > 0 Then
        ReDim Preserve aSID(k - 1)
        ReDim Preserve aUser(k - 1)
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
    
    Dim aRegRuns() As String, aDes() As String, result As SCAN_RESULT
    Dim i&, j&, sKey$, sData$, sHit$, sAlias$, sParam As String, sMD5$, aValue() As String
    Dim bData() As Byte, isDisabledWin8 As Boolean, isDisabledWinXP As Boolean, flagDisabled As Long, sKeyDisable As String
    Dim sFile$, sArgs$, sUser$, bSafe As Boolean, aLines() As String, sLine As String
    Dim aData() As String, bDisabled As Boolean, bMicrosoft As Boolean
    Dim sOrigLine As String
    
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
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL, HE_SID_ALL
    HE.AddKeys aRegRuns
    
    Do While HE.MoveNext
        
        
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
                            
                            If CBool(flagDisabled And 1) Then isDisabledWin8 = True
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
                
                sHit = sAlias & "[" & aValue(i) & "] = "
                
                sUser = ""
                If HE.IsSidUser Then
                    sUser = " (User '" & HE.UserName & "')"
                End If
                
                SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                
                sFile = FormatFileMissing(sFile, sArgs)
                
                bMicrosoft = False
                If InStr(1, sFile, "OneDrive", 1) <> 0 Then bMicrosoft = IsMicrosoftFile(sFile)
                
                sHit = sHit & ConcatFileArg(sFile, sArgs) & IIf(bMicrosoft, " (Microsoft)", "") & sUser
                bSafe = False
                
                If Not bIgnoreAllWhitelists And bHideMicrosoft Then
                    
                    '//TODO: narrow down to services' SID only: S-1-5-19 + S-1-5-20 + 'UpdatusUser' (NVIDIA)
                    
                    'Note: For services only
                    If StrComp(sFile, PF_64 & "\Windows Sidebar\Sidebar.exe", 1) = 0 And sArgs = "/autoRun" Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinDir & "\System32\mctadmin.exe", 1) = 0 And Len(sArgs) = 0 Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinSysDirWow64 & "\OneDriveSetup.exe", 1) = 0 And sArgs = "/thfirstsetup" Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    ElseIf StrComp(sFile, sWinDir & "\system32\SecurityHealthSystray.exe", 1) = 0 And Len(sArgs) = 0 Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                    
                    If OSver.MajorMinor = 5 And OSver.IsServer Then '2000 server
                        If aDes(HE.KeyIndex) = "RunOnce" Then
                            If WhiteListed(sFile, PF_32 & "\Internet Explorer\Connection Wizard\icwconn1.exe") And sArgs = "/desktop" Then bSafe = True
                        End If
                    End If
                    
                    If OSver.MajorMinor <= 6.1 Then 'Win2k-Win7
                        If WhiteListed(sFile, sWinDir & "\system32\CTFMON.EXE") And Len(sArgs) = 0 Then bSafe = True
                    End If

                    If OSver.MajorMinor = 6 Then 'Vista/2008
                        If WhiteListed(sFile, sWinDir & "\system32\rundll32.exe") And sArgs = "oobefldr.dll,ShowWelcomeCenter" Then
                            If IsMicrosoftFile(sWinDir & "\system32\oobefldr.dll") Then bSafe = True
                        End If
                    
                        If WhiteListed(sFile, PF_64 & "\Windows Sidebar\sidebar.exe") And (sArgs = "/autoRun" Or sArgs = "/detectMem") Then bSafe = True
                    End If
                    
                    If OSver.MajorMinor = 5 Then
                        If WhiteListed(sFile, sWinDir & "\system32\internat.exe") And Len(sArgs) = 0 Then bSafe = True
                    End If
                    
                    If OSver.MajorMinor >= 6.2 Then
                        If WhiteListed(sFile, PF_64 & "\Windows Defender\MSASCuiL.exe") And Len(sArgs) = 0 Then bSafe = True
                    End If

                End If
                
                If (Not bSafe) Or (Not bHideMicrosoft) Then
                
                    If (Not IsOnIgnoreList(sHit)) Then
                        
                        If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5
                        
                        With result
                            .Section = "O4"
                            .HitLineW = sHit
                            .Alias = sAlias
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i), , HE.Redirected
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                            
                            If Not isDisabledWinXP Then
                                AddProcessToFix .Process, FREEZE_OR_KILL_PROCESS, sFile
                            End If
                            .CureType = REGISTRY_BASED Or FILE_BASED Or PROCESS_BASED
                        End With
                        AddToScanResults result
                    End If
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
    aDes(1) = "Command Processor"
    
    aRegKey(2) = "HKLM\SYSTEM\CurrentControlSet\Control\BootVerificationProgram"
    aRegParam(2) = "ImagePath"
    aDefData(2) = ""
    aDes(2) = "BootVerificationProgram"
    
    aRegKey(3) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(3) = "BootExecute"
'    If OSver.MajorMinor = 5 Then 'Win2k
'        aDefData(3) = "autocheck autochk *" & vbNullChar & "DfsInit"
'    Else
'        aDefData(3) = "autocheck autochk *"
'    End If
    aDes(3) = "Session Manager"
    
    aRegKey(4) = "HKLM\SYSTEM\CurrentControlSet\Control\SafeBoot"
    aRegParam(4) = "AlternateShell"
    aDefData(4) = "cmd.exe"
    aDes(4) = "SafeBoot"
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        
        sParam = aRegParam(HE.KeyIndex)
        
        sData = Reg.GetData(HE.Hive, HE.Key, sParam, HE.Redirected)
        
        aData = SplitSafe(sData, vbNullChar) 'if MULTI_SZ (BootExecute)
        
        ArrayRemoveEmptyItems aData
        
        For i = 0 To UBound(aData)
        
            bSafe = False
        
            sData = aData(i)
            sOrigLine = sData
        
            If sParam = "BootExecute" Then
                If i = 0 Then
                    If StrBeginWith(sData, "autocheck ") Then 'remove autocheck, because it is not a real filename
                        sData = Mid$(sData, Len("autocheck ") + 1)
                    End If
                End If
                
                If bHideMicrosoft Then
                    If OSver.MajorMinor = 5 Then 'Win2k
                        If StrComp(sData, "autochk *", 1) = 0 Or StrComp(sData, "DfsInit", 1) = 0 Then bSafe = True
                    ElseIf OSver.MajorMinor >= 6.2 And OSver.IsServer Then '2012 Server, 2012 Server R2 (2016 too ?)
                        If StrComp(sData, "autochk /q /v *", 1) = 0 Then bSafe = True
                        If StrComp(sData, BuildPath(sWinSysDir, "autochk.exe") & " /q /v *", 1) = 0 Then bSafe = True
                    Else
                        If StrComp(sData, "autochk *", 1) = 0 Then bSafe = True
                    End If
                End If
            Else
                If sData = aDefData(HE.KeyIndex) Then bSafe = True
            End If
            
            bDisabled = False
            If sParam = "AlternateShell" Then
                If 1 <> Reg.GetDword(HKEY_LOCAL_MACHINE, HE.Key & "\Options", "UseAlternateShell") Then
                    bDisabled = True
                End If
            End If
            
            If Not bSafe Or bIgnoreAllWhitelists Or Not bHideMicrosoft Then
                
                'HKLM\..\Command Processor: [Autorun] =
                sAlias = IIf(bIsWin32, "O4", IIf(HE.Redirected, "O4-32", "O4")) & " - " & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": " & _
                    "[" & sParam & "] = "
                
                SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                
                sFile = FormatFileMissing(sFile)
                
                sHit = sAlias & ConcatFileArg(sFile, sArgs)
                
                If bDisabled Then sHit = sHit & " (disabled)"
                
                If Not IsOnIgnoreList(sHit) Then
                    
                    If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5
                    
                    With result
                        .Section = "O4"
                        .HitLineW = sHit
                        .Alias = sAlias
                        If sParam = "BootExecute" Then
                            
                            AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE, _
                                HE.Hive, HE.Key, sParam, , HE.Redirected, REG_RESTORE_MULTI_SZ, _
                                sOrigLine, "", vbNullChar
                            
                            AddRegToFix .Reg, APPEND_VALUE_NO_DOUBLE, HE.Hive, HE.Key, sParam, _
                                "autocheck autochk *", HE.Redirected, REG_RESTORE_MULTI_SZ
                            
                            If OSver.MajorMinor = 5 Then
                                AddRegToFix .Reg, APPEND_VALUE_NO_DOUBLE, HE.Hive, HE.Key, sParam, _
                                    "DfsInit", HE.Redirected, REG_RESTORE_MULTI_SZ
                            End If
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        Else
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile

                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End If
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    
    ReDim aRegKey(1 To 2) As String                   'key
    ReDim aRegParam(1 To UBound(aRegKey)) As String   'param
    ReDim aDes(1 To UBound(aRegKey)) As String        'description
    
  If bAdditional And Not bStartupScan Then
    
    'https://technet.microsoft.com/en-us/library/cc960241.aspx
    'PendingFileRenameOperations
    'Shared
    
    aRegKey(1) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(1) = "PendingFileRenameOperations"
    aDes(1) = "Session Manager"
    
    aRegKey(2) = "HKLM\System\CurrentControlSet\Control\Session Manager"
    aRegParam(2) = "PendingFileRenameOperations2"
    aDes(2) = "Session Manager"
    
    HE.Init HE_HIVE_HKLM, , HE_REDIR_NO_WOW
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        
        sParam = aRegParam(HE.KeyIndex)
    
        sData = Reg.GetData(HE.Hive, HE.Key, sParam, HE.Redirected)
        
        If Len(sData) <> 0 Then
        
          'converting MULTI_SZ to [1] -> [2], [3] -> [4] ...
          aLines = SplitSafe(sData, vbNullChar)
        
          For j = 0 To UBound(aLines) Step 2
            sFile = PathNormalize(aLines(j))
            If j + 1 <= UBound(aLines) Then
                If aLines(j + 1) = "" Then
                    sArgs = "-> DELETE"
                Else
                    sArgs = "-> " & PathNormalize(aLines(j + 1))
                End If
            End If
            
            'HKLM\..\FileRenameOperations:
            sAlias = IIf(bIsWin32, "O4", IIf(HE.Redirected, "O4-32", "O4")) & " - " & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": " & _
                "[" & sParam & "] = "
            
            sFile = FormatFileMissing(sFile)
            
            sHit = sAlias & ConcatFileArg(sFile, sArgs)
            
            'If Not IsOnIgnoreList(sHit) And Not FileMissing(sFile) Then
            
            If Not IsOnIgnoreList(sHit) Then
            
              'If sArgs <> "-> DELETE" Or (sArgs = "-> DELETE" And bShowPendingDeleted) Then

                If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5

                With result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sParam, , HE.Redirected
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                    
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults result
              'End If
            End If
          Next
        End If
    Loop
    
  End If
    
    Dim aFiles() As String
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
                    If (StrBeginWith(aData(j), "REM") And Not bIgnoreAllWhitelists) Then
                        aData(j) = ""
                    ElseIf (StrBeginWith(aData(j), "::") And Not bIgnoreAllWhitelists) Then
                        aData(j) = ""
                    ElseIf bHideMicrosoft Then
                        'check whitelist
                        If StrEndWith(sFile, "AutoExec.nt") Then
                            If StrComp(aData(j), "@echo off", 1) = 0 Then
                                aData(j) = ""
                            ElseIf StrComp(aData(j), "lh %SystemRoot%\system32\mscdexnt.exe", 1) = 0 Then
                                aData(j) = ""
                            ElseIf StrComp(aData(j), "lh %SystemRoot%\system32\redir", 1) = 0 Then
                                aData(j) = ""
                            ElseIf StrComp(aData(j), "lh %SystemRoot%\system32\dosx", 1) = 0 Then
                                aData(j) = ""
                            ElseIf StrComp(aData(j), "SET BLASTER=A220 I5 D1 P330 T3", 1) = 0 Then
                                aData(j) = ""
                            End If
                        ElseIf StrEndWith(sFile, "Config.nt") Then
                            If StrComp(aData(j), "dos=high, umb", 1) = 0 Then
                                aData(j) = ""
                            ElseIf StrComp(aData(j), "device=%SystemRoot%\system32\himem.sys", 1) = 0 Then
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
                            If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5
                            With result
                                .Section = "O4"
                                .HitLineW = sHit
                                .Alias = sAlias
                                AddFileToFix .File, REMOVE_FILE, sFile
                                .CureType = FILE_BASED
                            End With
                            AddToScanResults result
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
    Dim pos As Long
    
    aRegKey(1) = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
    aDes(1) = "RunOnceEx"
    aRegKey(2) = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnceEx"
    aDes(2) = "RunServicesOnceEx"
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys aRegKey
    
    Do While HE.MoveNext
        If Reg.KeyHasSubKeys(HE.Hive, HE.Key, HE.Redirected) Then
            
            For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey(), HE.Redirected, , False)
                
                For j = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key & "\" & aSubKey(i), aValue(), HE.Redirected)
                    
                    sData = Reg.GetString(HE.Hive, HE.Key & "\" & aSubKey(i), aValue(j), HE.Redirected)
                    
                    'e.g. C:\PROGRA~2\COMMON~1\MICROS~1\Repostry\REPCDLG.OCX|DllRegisterServer
                    pos = InStr(sData, "|")
                    If pos <> 0 Then
                        sFile = Left$(sData, pos - 1)
                        sArgs = Mid$(sData, pos)
                    Else
                        sFile = sData
                        sArgs = vbNullString
                    End If
                    
                    sFile = FormatFileMissing(sFile)
                    
                    sAlias = "O4 - " & HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & ": "
                    sHit = sAlias & aSubKey(i) & " [" & aValue(j) & "] = " & ConcatFileArg(sFile, sArgs)
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O4"
                            .HitLineW = sHit
                            .Alias = sAlias
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key & "\" & aSubKey(i), aValue(j), , HE.Redirected
                            AddRegToFix .Reg, REMOVE_KEY_IF_NO_VALUES, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                Next
            Next
        End If
    Loop
    
  If bAdditional Then
    
    'Autorun.inf
    'http://journeyintoir.blogspot.com/2011/01/autoplay-and-autorun-exploit-artifacts.html
    Dim aDrives() As String
    Dim sAutorun As String
    Dim aVerb() As String
    Dim bOnce As Boolean

    aVerb = Split("open|shellexecute|shell\open\command|shell\explore\command", "|")

    ' Mapping scheme for "inf. verb" -> to "registry" :
    '
    ' icon                  -> _Autorun\Defaulticon
    ' open                  -> shell\AutoRun\command
    ' shellexecute          -> shell\AutoRun\command
    ' shell\open\command    -> shell\open\command
    ' shell\explore\command -> shell\explore\command

    aDrives = GetDrives(DRIVE_BIT_FIXED Or DRIVE_BIT_REMOVABLE)

    For i = 1 To UBound(aDrives)
        sAutorun = BuildPath(aDrives(i), "autorun.inf")
        If FileExists(sAutorun) Then

            bOnce = False

            For j = 0 To UBound(aVerb)

                sFile = ""
                sArgs = ""
                sData = ReadIniA(sAutorun, "autorun", aVerb(j))

                If Len(sData) <> 0 Then
                    SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=False
                    sFile = FormatFileMissing(sFile)

                    sHit = "O4 - Autorun.inf: " & sAutorun & " - " & aVerb(j) & " - " & ConcatFileArg(sFile, sArgs)

                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O4"
                            .HitLineW = sHit
                            AddFileToFix .File, REMOVE_FILE, sAutorun
                            .CureType = FILE_BASED
                        End With
                        AddToScanResults result
                    End If

                    bOnce = True
                End If
            Next

            'if unknown data is inside autorun.inf
            If Not bOnce Then

                sHit = "O4 - Autorun.inf: " & sAutorun & " - " & "(unknown target)"

                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O4"
                        .HitLineW = sHit
                        AddFileToFix .File, REMOVE_FILE, sAutorun
                        .CureType = FILE_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        End If
    Next
    
  End If
  
  If bAdditional Then

    'MountPoints2
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"

    aVerb = Split("shell\AutoRun\command|shell\open\command|shell\explore\command", "|")

    Do While HE.MoveNext
        
        For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey, HE.Redirected)
            For j = 0 To UBound(aVerb)
                sKey = HE.Key & "\" & aSubKey(i) & "\" & aVerb(j)

                If Reg.KeyExists(HE.Hive, sKey, HE.Redirected) Then

                    sData = Reg.GetString(HE.Hive, sKey, "", HE.Redirected)

                    SplitIntoPathAndArgs sData, sFile, sArgs, bIsRegistryData:=True
                    sFile = FormatFileMissing(sFile)

                    sHit = IIf(HE.Redirected, "O4-32", "O4") & " - MountPoints2: " & HE.HiveNameAndSID & "\..\" & aSubKey(i) & "\" & aVerb(j) & ": (default) = " & ConcatFileArg(sFile, sArgs)

                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O4"
                            .HitLineW = sHit
                            'remove MountPoints2\{CLSID}
                            'or
                            'remove MountPoints2\Letter
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        Next
    Loop
    
  End If
    
    'ScreenSaver
    
    HE.Init HE_HIVE_HKCU Or HE_HIVE_HKU, HE_SID_USER Or HE_SID_NO_VIRTUAL, HE_REDIR_NO_WOW
    HE.AddKey "Control Panel\Desktop"
    
    Do While HE.MoveNext
    
      sFile = Reg.GetString(HE.Hive, HE.Key, "SCRNSAVE.EXE")
      If 0 <> Len(sFile) Then
        bSafe = True
        sOrigLine = sFile

        sFile = FormatFileMissing(sFile)
        
        If FileMissing(sFile) Then
            '(Нет)
            If (sOrigLine <> STR_CONST.RU_NO And sOrigLine <> "(None)" Or bIgnoreAllWhitelists) Then
                bSafe = False
            End If
        Else
            If Not IsMicrosoftFile(sFile) Then bSafe = False
        End If
        
        If Not bSafe Then
            sHit = "O4 - " & HE.HiveNameAndSID & "\Control Panel\Desktop: [SCRNSAVE.EXE] = " & sFile
        
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O4"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "SCRNSAVE.EXE"
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults result
            End If
        End If
      End If
    Loop
    
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
    
    Dim sHive$, i&, j&, sAlias$, sMD5$, result As SCAN_RESULT
    Dim aSubKey$(), sDay$, sMonth$, sYear$, sKey$, sFile$, sTime$, sHit$, SourceHive$, dEpoch As Date, sArgs$, sUser$, sDate$
    Dim Values$(), bData() As Byte, flagDisabled As Long, dDate As Date, UseWow As Variant, Wow6432Redir As Boolean, sTarget$, sData$
    Dim bMicrosoft As Boolean
    
    Const sDateEpoch As String = "1601/01/01"
    
    If OSver.MajorMinor >= 6.2 Then ' Win 8+
    
        For i = 0 To UBound(aHives) 'HKLM, HKCU, HKU\SID()

            sHive = aHives(i)
            
            For Each UseWow In Array(False, True)
    
                Wow6432Redir = UseWow
  
                If (bIsWin32 And Wow6432Redir) _
                  Or bIsWin64 And Wow6432Redir And (sHive = "HKCU" Or StrBeginWith(sHive, "HKU\")) Then
                    Exit For
                End If
            
                
                For j = 1 To Reg.EnumValuesToArray(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values())
            
                    flagDisabled = 2
                    ReDim bData(0)
                    
                    bData() = Reg.GetBinary(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values(j))
                    
                    'undoc. flag. Seen:
                    '0x02000000 - enabled
                    '0x06000000 - enabled
                    '0x03000000 - disabled
                    '0x07000000 - disabled
                    '---
                    'looks like, flag 1 - is disabled
                    If UBoundSafe(bData) >= 11 Then
                        GetMem4 ByVal VarPtr(bData(0)), flagDisabled
                    End If
                    
                    If AryItems(bData) And CBool(flagDisabled And 1) Then   'is Disabled ?
                    
                        dDate = ConvertFileTimeToLocalDate(VarPtr(bData(4)))
                        
                        If Reg.ValueExists(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), Wow6432Redir) Then
                        
                            sData = Reg.GetString(0&, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), Wow6432Redir)
                        
                            'if you change it, change fix appropriate !!!
                            sAlias = "O4 - " & sHive & "\..\StartupApproved\" & IIf(bIsWin64 And Wow6432Redir, "Run32", "Run") & ": "
                            
                            sHit = sAlias & "[" & Values(j) & "] "
                            
                            sUser = ""
                            If aUser(i) <> "" And StrBeginWith(sHive, "HKU\") Then
                                If (sHive <> "HKU\S-1-5-18" And _
                                    sHive <> "HKU\S-1-5-19" And _
                                    sHive <> "HKU\S-1-5-20") Then sUser = " (User '" & aUser(i) & "')"
                            End If
                            
                            SplitIntoPathAndArgs sData, sFile, sArgs, True
                            
                            sFile = FormatFileMissing(sFile)
                            
                            bMicrosoft = False
                            If InStr(1, sFile, "Windows Defender", 1) <> 0 Then bMicrosoft = IsMicrosoftFile(sFile)
                            
                            sHit = sHit & "= " & ConcatFileArg(sFile, sArgs) & IIf(bMicrosoft, " (Microsoft)", "") & sUser
                            
                            sDate = Format$(dDate, "yyyy\/mm\/dd")
                            If sDate <> sDateEpoch Then sHit = sHit & " (" & sDate & ")"
                            
                            If Not IsOnIgnoreList(sHit) Then
                            
                                If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5
                
                                With result
                                    .Section = "O4"
                                    .HitLineW = sHit
                                    .Alias = sAlias
                                    AddRegToFix .Reg, REMOVE_VALUE, 0, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\" & _
                                        IIf(bIsWin64 And Wow6432Redir, "Run32", "Run"), Values(j), , False
                                    AddRegToFix .Reg, REMOVE_VALUE, 0, sHive & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Values(j), , CLng(Wow6432Redir)
                                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                                    .CureType = REGISTRY_BASED Or FILE_BASED
                                End With
                                AddToScanResults result
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
            
            sHit = sAlias & aSubKey(i) & " [command] = "

            SplitIntoPathAndArgs sData, sFile, sArgs, True
            
            sFile = FormatFileMissing(sFile)
            
            If SourceHive <> "" Then sArgs = sArgs & IIf(Len(sArgs) = 0, "", " ") & "(" & SourceHive & ")"
            sArgs = sArgs & " (" & sTime & ")"
            
            sHit = sHit & ConcatFileArg(sFile, sArgs)
           
            If Not IsOnIgnoreList(sHit) Then
                
                If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5
                
                With result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_KEY, 0, sKey & "\" & aSubKey(i), , , REG_NOTREDIRECTED
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        Next
        
        'Startup folder items
        
        sKey = "HKLM\SOFTWARE\Microsoft\Shared Tools\MSConfig\startupfolder"
        
        
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
                sHit = sAlias & aSubKey(i) & " [backup] => " & sTarget & IIf(sArgs <> "", " " & sArgs, "") & " (" & sTime & ")" & IIf(Not FileExists(sTarget), " (file missing)", "")
            Else
                sHit = sAlias & aSubKey(i) & " [backup] = " & sFile & " (" & sTime & ")" & IIf(sFile = "", " (no file)", IIf(Not FileExists(sFile), " (file missing)", ""))
            End If
            
            If Not IsOnIgnoreList(sHit) Then
                
                If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5
                
                With result
                    .Section = "O4"
                    .HitLineW = sHit
                    .Alias = sAlias
                    AddRegToFix .Reg, REMOVE_KEY, 0&, sKey & "\" & aSubKey(i), , , REG_NOTREDIRECTED
                    AddFileToFix .File, REMOVE_FILE, sFile 'removing backup (.pss)
                    .CureType = FILE_BASED Or REGISTRY_BASED
                End With
                AddToScanResults result
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
    
    Dim aRegKeys() As String, aParams() As String, aDes() As String, aDesConst() As String, result As SCAN_RESULT
    Dim sAutostartFolder$(), sShortCut$, i&, k&, Wow6432Redir As Boolean, UseWow, sFolder$, sHit$, dEpoch As Date
    Dim FldCnt&, sKey$, sSID$, sFile$, sLinkPath$, sLinkExt$, sTarget$, Blink As Boolean, bPE_EXE As Boolean
    Dim bData() As Byte, isDisabled As Boolean, flagDisabled As Long, sKeyDisable As String, sHive As String, dDate As Date
    Dim StartupCU As String, aFiles() As String, sArguments As String, aUserNames() As String, aUserConst() As String, sUsername$
    Dim aFolders() As String
    
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
    For k = 1 To UBound(aRegKeys)
    
        For Each UseWow In Array(False, True)
            
            Wow6432Redir = UseWow
        
            'skip HKCU Wow64
            If (bIsWin32 And Wow6432Redir) _
              Or bIsWin64 And Wow6432Redir And StrBeginWith(aRegKeys(k), "HKCU") Then Exit For
    
            FldCnt = FldCnt + 1
            sAutostartFolder(FldCnt) = Reg.GetString(0&, aRegKeys(k), aParams(k), Wow6432Redir)
            aDes(FldCnt) = aDesConst(k)
            aUserNames(FldCnt) = aUserConst(k)
            
            'save path of Startup for current user to substitute other user names
            If aParams(k) = "Startup" Then
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
            
            For k = 1 To UBound(aRegKeys)
            
                'only HKCU keys
                If StrBeginWith(aRegKeys(k), "HKCU") Then
                
                    ' Convert HKCU -> HKU
                    sKey = Replace$(aRegKeys(k), "HKCU\", "HKU\" & sSID)
                
                    FldCnt = FldCnt + 1
                    If UBound(sAutostartFolder) < FldCnt Then
                        ReDim Preserve sAutostartFolder(UBound(sAutostartFolder) + 100)
                        ReDim Preserve aDes(UBound(aDes) + 100)
                        ReDim Preserve aUserNames(UBound(aUserNames) + 100)
                    End If
            
                    sAutostartFolder(FldCnt) = Reg.GetString(0&, sKey, aParams(k))
                    aDes(FldCnt) = sSID & " " & aDesConst(k)
                    aUserNames(FldCnt) = aUser(i)
                End If
            Next
        End If
    Next
    
    ReDim Preserve sAutostartFolder(FldCnt)
    ReDim Preserve aDes(FldCnt)
    ReDim Preserve aUserNames(FldCnt)
    
    For k = 1 To UBound(sAutostartFolder)
        sAutostartFolder(k) = UnQuote(EnvironW(sAutostartFolder(k)))
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
    
    For k = 1 To UBound(sAutostartFolder)
        
        sUsername = aUserNames(k)
        
        sFolder = sAutostartFolder(k)
        
        If 0 <> Len(sFolder) Then
          If FolderExists(sFolder) Then
            
            Erase aFolders
            aFolders = ListSubfolders(sFolder)
            
            For i = 0 To UBoundSafe(aFolders)
            
                sHit = "O4 - " & aDes(k) & ": " & aFolders(i) & " (folder)"
            
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O4"
                        .HitLineW = sHit
                        AddFileToFix .File, REMOVE_FOLDER, aFolders(i)
                        .CureType = FILE_BASED
                    End With
                    AddToScanResults result
                End If
            Next
            
            Erase aFiles
            aFiles = ListFiles(sFolder)
            
            For i = 0 To UBoundSafe(aFiles)
            
                sShortCut = GetFileNameAndExt(aFiles(i))

                If (LCase$(sShortCut) <> "desktop.ini" Or bIgnoreAllWhitelists) Then

                  If Not FolderExists(sFolder & "\" & sShortCut) Then
                  
                    isDisabled = False
              
                    If OSver.MajorMinor >= 6.2 Then  ' Win 8+

                        If StrInParamArray(aDes(k), "Startup", "User Startup", "Global Startup", "Global User Startup") Then

                            Select Case aDes(k)
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
                        sHit = "O4 - " & aDes(k) & ": "
                    End If
                    
                    If StrInParamArray(sLinkExt, ".LNK", ".URL", ".WEBSITE", ".PIF") Then Blink = True
                    
                    If Not Blink Or sLinkExt = ".PIF" Then  'not a Shortcut ?
                        bPE_EXE = isPE(sLinkPath)       'PE EXE ?
                    End If
                    
                    sTarget = ""
                    
                    If Blink Then
                        sTarget = GetFileFromShortcut(sLinkPath, sArguments)
                            
                        sHit = sHit & aFiles(i) & "    ->    " & sTarget & IIf(Len(sArguments) <> 0, " " & sArguments, "") 'doSafeURLPrefix
                    Else
                        sHit = sHit & aFiles(i) & IIf(bPE_EXE, "    ->    (PE EXE)", "")
                    End If
                    
                    'If sUsername <> "" Then sHit = sHit & " (Folder '" & sUsername & "')"
                    
                    If isDisabled Then sHit = sHit & IIf(dDate <> dEpoch, " (" & Format$(dDate, "yyyy\/mm\/dd") & ")", "")
                    
                    If Not IsOnIgnoreList(sHit) Then
                        
                        If g_bCheckSum Then
                            If Not Blink Or bPE_EXE Then
                                sHit = sHit & GetFileCheckSum(sLinkPath)
                            Else
                                If 0 <> Len(sTarget) Then
                                    sHit = sHit & GetFileCheckSum(sTarget)
                                End If
                            End If
                        End If
                        
                        With result
                          .Section = "O4"
                          .HitLineW = sHit
                          
                          If isDisabled Then
                            .Alias = sHive & "\..\StartupApproved\StartupFolder:"
                            AddRegToFix .Reg, REMOVE_VALUE, 0&, sKeyDisable, sShortCut, , REG_NOTREDIRECTED
                            AddFileToFix .File, REMOVE_FILE, sLinkPath
                            .CureType = FILE_BASED Or REGISTRY_BASED
                          Else
                            .Alias = aDes(k)
                            AddFileToFix .File, REMOVE_FILE, sLinkPath
                            AddProcessToFix .Process, FREEZE_OR_KILL_PROCESS, sTarget
                            If Blink Then
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sTarget
                            End If
                            .CureType = FILE_BASED Or PROCESS_BASED
                          End If
                        End With
                        AddToScanResults result
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
    '2.7.0.26 - Added scanning for Folders in AutoStart folder locations.
    
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
    
    Call Reg.EnumSubKeysToArray(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", gSID_All)
    
    'Note:
    '
    'gSIDs - include all active SIDs, excluding current user
    'gSID_All - include all active and non-active SIDs with current user as well
    
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
    gUsers(UBound(gHives)) = OSver.UserName
    
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

Public Sub FixO4Item(sItem$, result As SCAN_RESULT)
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

    FixProcessHandler result
    
    If InStr(sItem, "StartupApproved\StartupFolder") <> 0 Then
        
        sFile = result.File(0).Path
        
        If FileExists(sFile) Then
            If DeleteFileForce(sFile) Then
                FixRegistryHandler result 'remove registry value if only file successfully deleted (!!!)
            End If
        Else
            FixRegistryHandler result
        End If
        
        Exit Sub
    End If
    
    FixFileHandler result
    FixRegistryHandler result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO4Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO5Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO5Item - Begin"
    
    Dim sControlIni$, sDummy$, sHit$, result As SCAN_RESULT
    Dim i&, aValues() As String, bSafe As Boolean, bFileExist As Boolean, sSnapIn As String, sDescr As String, sPath As String
    Dim aParams() As Variant
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    '// TODO: add also:
    
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => DisallowCpl = 1
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer => RestrictCpl = 1
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "HKCU\Control Panel\don't load"
    HE.AddKey "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\don't load"
            
    Do While HE.MoveNext
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValues, HE.Redirected)
            
            bSafe = False
            sSnapIn = aValues(i)
            
            If bHideMicrosoft Then
                If HE.Hive = HKCU Or HE.Hive = HKU Then
                    If inArraySerialized(sSnapIn, sSafeO5Items_HKU, "|", , , 1) Then bSafe = True
                Else
                    If HE.Redirected Then
                        If inArraySerialized(sSnapIn, sSafeO5Items_HKLM_32, "|", , , 1) Then bSafe = True
                    Else
                        If inArraySerialized(sSnapIn, sSafeO5Items_HKLM, "|", , , 1) Then bSafe = True
                    End If
                End If
            End If
            
            If Not bSafe Then
                sPath = BuildPath(IIf(HE.Redirected, sWinSysDirWow64, sWinSysDir), sSnapIn)
                bFileExist = FileExists(sPath)
                sDescr = ""
                If bFileExist Then
                    sDescr = GetFileProperty(sPath, "FileDescription")
                End If
                
                sHit = IIf(HE.Redirected, "O5-32", "O5") & " - " & HE.KeyAndHive & ": [" & sSnapIn & "]" & _
                    IIf(Len(sDescr) <> 0, " (" & sDescr & ")", "") & IIf(bFileExist, "", " (file missing)")
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O5"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sSnapIn, , HE.Redirected
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    sControlIni = sWinDir & "\control.ini"
    If Not FileExists(sControlIni) Then Exit Sub
    
    Dim cIni As clsIniFile
    Set cIni = New clsIniFile
    
    cIni.InitFile sControlIni, 1251
    
    If cIni.CountParams("don't load") > 0 Then
        aParams = cIni.GetParamNames("don't load")
        
        For i = 0 To UBound(aParams)
            sSnapIn = aParams(i)
            sDummy = Trim$(cIni.ReadParam("don't load", sSnapIn))
            
            If Len(sDummy) <> 0 Then
                sPath = BuildPath(IIf(HE.Redirected, sWinSysDirWow64, sWinSysDir), sSnapIn)
                bFileExist = FileExists(sPath)
                sDescr = ""
                If bFileExist Then
                    sDescr = GetFileProperty(sPath, "FileDescription")
                End If
                
                sHit = "O5 - control.ini: [don't load] " & sSnapIn & " = " & sDummy & _
                    IIf(Len(sDescr) <> 0, " (" & sDescr & ")", "") & IIf(bFileExist, "", " (file missing)")
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O5"
                        .HitLineW = sHit
                        AddIniToFix .Reg, RESTORE_VALUE_INI, "control.ini", "don't load", sSnapIn, vbNullString
                        .CureType = INI_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    End If

    Set cIni = Nothing
    'Set HE = Nothing
    
    AppendErrorLogCustom "CheckO5Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO5Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO5Item(sItem$, result As SCAN_RESULT)
    'O5 - Blocking of loading Internet Options in Control Panel
    'WritePrivateProfileString "don't load", "inetcpl.cpl", vbNullString, "control.ini"
    On Error GoTo ErrorHandler:
    FixRegistryHandler result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO5Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO6Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO6Item - Begin"
    
    'If there are sub folders called
    '"restrictions" and/or "control panel", delete them
    
    Dim sHit$, Key$(2), result As SCAN_RESULT
    'keys 0,1,2 - are x6432 shared.
    
    Key(0) = "Software\Policies\Microsoft\Internet Explorer\Restrictions"
    Key(1) = "Software\Policies\Microsoft\Internet Explorer\Toolbars\Restrictions"
    Key(2) = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys Key()
    
    Do While HE.MoveNext
        If Reg.KeyHasValues(HE.Hive, HE.Key, HE.Redirected) Then
            sHit = IIf(HE.Redirected, "O6-32", "O6") & " - IE Policy: " & HE.HiveNameAndSID & "\" & HE.Key & " - present"
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O6"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key, , , HE.Redirected
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Loop
    
    'Restrict using IE settings by HKLM hive only ?
    If Reg.GetDword(HKLM, "SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings", "Security_HKLM_only") = 1 Then
        sHit = "O6 - IE Policy: HKLM\..\Internet Settings: [Security_HKLM_only] = 1"
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O6"
                .HitLineW = sHit
                AddRegToFix .Reg, REMOVE_VALUE, HKLM, "SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings", "Security_HKLM_only"
                .CureType = REGISTRY_BASED
            End With
            AddToScanResults result
        End If
    End If
    
    AppendErrorLogCustom "CheckO6Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO6Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO6Item(sItem$, result As SCAN_RESULT)
    'O6 - Disabling of Internet Options' Main tab with Policies
    FixRegistryHandler result
End Sub

Public Sub CheckSystemProblems()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckSystemProblems - Begin"
    
    Call CheckSystemProblemsEnvVars
    Call CheckSystemProblemsFreeSpace
    Call CheckSystemProblemsNetwork
    
    AppendErrorLogCustom "CheckSystemProblems - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblems"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckSystemProblemsEnvVars()
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "CheckSystemProblemsEnvVars - Begin"
    
    'Checking for present and correct type of parameters:
    'HKCU\Environment => temp, tmp
    '+HKU
    
    'Checking for present, correct type of parameters and correct value:
    'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment => temp, tmp ("%SystemRoot%\TEMP")
    
    Dim sData As String, sDataNonExpanded As String
    Dim vParam, sKeyFull As String, sHit As String, sDefValue As String, result As SCAN_RESULT
    Dim aLine() As String, i As Long, vValue, bSafe As Boolean, sPsPath As String
    Dim bComply As Boolean
    
    '// TODO:
    ' PATH len exceed the maximum allowed, see article:
    ' https://safezone.cc/threads/delo-o-zablokirovannoj-peremennoj-okruzhenija-path.31001/
    
    ' Check essential programs, e.g. scripting hosts, by search path.
    
    ' Add new field to 'Result' - reboot required
    
    ' %PATH% - Check for system folders presence
    ' Add env. vars (partial log) in case problems with %PATH%
    
    sKeyFull = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    
    sData = Reg.GetString(0, sKeyFull, "Path")
    aLine = Split(sData, ";")
    
    For i = 0 To UBoundSafe(aLine)
        aLine(i) = LTrim$(aLine(i))
        If StrEndWith(aLine(i), "\") Then aLine(i) = Left$(aLine(i), Len(aLine(i)) - 1) 'cut last \
    Next
    
    If OSver.IsWindows7OrGreater Then sPsPath = BuildPath(sWinSysDir, "WindowsPowerShell\v1.0")
    
    For Each vValue In Array(sWinDir, sWinSysDir, BuildPath(sWinSysDir, "Wbem"), sPsPath)
        
        bSafe = False
        
        If Len(vValue) = 0 Then
            bSafe = True
        Else
            If AryItems(aLine) Then
                If inArray(CStr(vValue), aLine, , , vbTextCompare) Then bSafe = True
            End If
        End If
        
        If Not bSafe Then
            sHit = "O7 - TroubleShooting: (EV) %PATH% has missing system folder: " & vValue
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, APPEND_VALUE_NO_DOUBLE, 0, sKeyFull, "Path", EnvironUnexpand(CStr(vValue)) & ";", , REG_RESTORE_EXPAND_SZ
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    '// TODO:
    'add checking the popular exe-files + ext. that hijack %PATH% of the normal similar exe names
    
    'Check for %Temp% anomalies
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    If OSver.IsWindows7OrGreater Then
        '7+
        HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    Else
        'Vista-
        HE.Init HE_HIVE_ALL, (HE_SID_ALL And Not HE_SID_SERVICE) Or HE_SID_NO_VIRTUAL, HE_REDIR_NO_WOW
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
                
                bComply = False
                
                If OSver.IsElevated Then
                    bComply = True
                Else
                    'these keys are access restricted for Limited users
                    If Not (HE.Hive = HKU And (HE.SID = "S-1-5-19" Or HE.SID = "S-1-5-20")) Then
                        bComply = True
                    End If
                End If
                
                If bComply Then
                    sHit = "O7 - TroubleShooting: (EV) " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = (not exist)"
                End If
            Else
                sData = Reg.GetString(0, sKeyFull, CStr(vParam))
                sDataNonExpanded = Reg.GetString(0, sKeyFull, CStr(vParam), , True)
                
                If InStr(sData, "%") <> 0 Then
                    sHit = "O7 - TroubleShooting: (EV) " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = " & sData & " (wrong type of parameter)"
                ElseIf sData = "" Then
                    sHit = "O7 - TroubleShooting: (EV) " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = (empty value)"
                End If
                
                sData = EnvironW(sData)
                
                If sHit = "" Then
                    If Not FolderExists(sData) Then
                        sHit = "O7 - TroubleShooting: (EV) " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = " & sData & " (folder missing)"
                    End If
                    
'                    If HE.Hive = HKLM Then
'                        If StrComp(sData, SysDisk & "\TEMP", 1) <> 0 _
'                          And StrComp(sData, sWinDir & "\TEMP", 1) <> 0 Then 'if wrong value
'                            sHit = "O7 - TroubleShooting: [EV] " & HE.HiveNameAndSID & "\..\Environment: " & "[" & vParam & "]" & " = " & sData & " (environment value is altered)"
'                        End If
'                    Else
'                        If OSver.MajorMinor < 6 Then
'                            If StrComp(sData, UserProfile & "\Local Settings\Temp", 1) <> 0 _
'                              And StrComp(sData, SysDisk & "\TEMP", 1) <> 0 Then 'if wrong value
'                                sHit = "O7 - TroubleShooting: [EV] " & HE.HiveNameAndSID & _
'                                  "\..\Environment: " & "[" & vParam & "]" & " = " & sData & " (environment value is altered)"
'                            End If
'                        Else
'                            If StrComp(sData, LocalAppData & "\Temp", 1) <> 0 _
'                              And StrComp(sData, SysDisk & "\TEMP", 1) <> 0 Then 'if wrong value
'                                sHit = "O7 - TroubleShooting: [EV] " & HE.HiveNameAndSID & _
'                                  "\..\Environment: " & "[" & vParam & "]" & " = " & sData & " (environment value is altered)"
'                            End If
'                        End If
'                    End If

                End If
            End If
            
            If sHit <> "" Then
                If Not IsOnIgnoreList(sHit) Then
                    With result
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
                        
                        If StrEndWith(sHit, "(folder missing)") Then
                            AddFileToFix .File, CREATE_FOLDER, sData
                            .CureType = .CureType Or FILE_BASED
                        End If
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    AppendErrorLogCustom "CheckSystemProblemsEnvVars - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblemsEnvVars"
    If inIDE Then Stop: Resume Next
End Sub
    
Public Sub CheckSystemProblemsFreeSpace()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckSystemProblemsFreeSpace - Begin"
    
    Dim cFreeSpace As Currency
    Dim sHit As String
    Dim result As SCAN_RESULT
    
    cFreeSpace = GetFreeDiscSpace(SysDisk, False)
    ' < 1 GB ?
    If (cFreeSpace < cMath.MBToInt64(1& * 1024)) And (cFreeSpace <> 0@) Then
        
        sHit = "O7 - TroubleShooting: (Disk) Free disk space on " & SysDisk & " is too low = " & (cFreeSpace / 1024& / 1024& * 10000& \ 1) & " MB."
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O7"
                .HitLineW = sHit
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
    End If
    
    AppendErrorLogCustom "CheckSystemProblemsFreeSpace - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblemsFreeSpace"
    If inIDE Then Stop: Resume Next
End Sub
    
Public Sub CheckSystemProblemsNetwork()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckSystemProblemsNetwork - Begin"
    
    Dim sNetBiosName As String
    Dim sHit As String
    Dim result As SCAN_RESULT
    
    If GetCompName(ComputerNamePhysicalDnsHostname) = "" Then
    
        sNetBiosName = GetCompName(ComputerNameNetBIOS)
        sHit = "O7 - TroubleShooting: (Network) Computer name (hostname) is not set" & IIf(sNetBiosName <> "", " (should be: " & sNetBiosName & ")", "")
        
        If Not IsOnIgnoreList(sHit) Then
            With result
                .Section = "O7"
                .HitLineW = sHit
                .CureType = CUSTOM_BASED
            End With
            AddToScanResults result
        End If
    End If
    
    AppendErrorLogCustom "CheckSystemProblemsNetwork - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckSystemProblemsNetwork"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckCertificatesEDS()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckCertificatesEDS - Begin"
    
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
    
    Dim i&, aSubKey$(), idx&, sTitle$, bSafe As Boolean, sHit$, result As SCAN_RESULT, ResultAll As SCAN_RESULT
    Dim Blob() As Byte, CertHash As String, FriendlyName As String, IssuedTo As String, nItems As Long
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\SystemCertificates\Disallowed\Certificates"
    
    Do While HE.MoveNext
        For i = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey())
            
            bSafe = True
            sTitle = ""
            
            Blob = Reg.GetBinary(HE.Hive, HE.Key & "\" & aSubKey(i), "Blob")
            
            If AryItems(Blob) Then
                ParseCertBlob Blob, CertHash, FriendlyName, IssuedTo
                
                If CertHash = "" Then CertHash = aSubKey(i)
                
                idx = GetCollectionIndexByKey(CertHash, colDisallowedCert)
                
                If idx <> 0 Then
                    'it's safe
                    If Not bHideMicrosoft Or bIgnoreAllWhitelists Then
                        sTitle = colDisallowedCert(idx)
                        bSafe = False
                    End If
                Else
                    bSafe = False
                End If
                
                If Not bSafe Then
                    If sTitle = "" Then sTitle = IssuedTo
                    If sTitle = "" Then sTitle = "Unknown"
                    If FriendlyName <> "" Then sTitle = sTitle & " (" & FriendlyName & ")"
                    If FriendlyName = "Fraudulent" Or FriendlyName = "Untrusted" Then sTitle = sTitle & " (HJT: possible, safe)"
                    
                    'O7 - Policy: [Untrusted Certificate] Hash - 'Name, cert. issued to' (HJT rating, if possible)
                    sHit = "O7 - Policy: [Untrusted Certificate] " & HE.HiveNameAndSID & " - " & CertHash & " - " & sTitle
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i)
                            AddRegToFix ResultAll.Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                        nItems = nItems + 1
                    End If
                End If
            End If
        Next
    Loop
    
    If nItems > 10 Then
        With ResultAll
            .Alias = "O7"
            .HitLineW = "O7 - Policy: [Untrusted Certificate] Fix all items from the log"
            .CureType = REGISTRY_BASED
        End With
        AddToScanResults ResultAll
    End If
    
    'Check for new Microsoft Root certificates
    Dim sData$
    Dim eHive As Long, vHive As Variant
    
    For Each vHive In Array(HKCU, HKLM)
        eHive = vHive
    
        For i = 1 To Reg.EnumSubKeysToArray(eHive, "SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates", aSubKey())
            
            If Not (IsMicrosoftCertHash(aSubKey(i))) Then
            
                Blob = Reg.GetBinary(eHive, "SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates\" & aSubKey(i), "Blob")
                
                If AryItems(Blob) Then
                    ParseCertBlob Blob, CertHash, FriendlyName, IssuedTo
                    
                    If InStr(1, FriendlyName, "Microsoft", 1) <> 0 And IssuedTo <> "localhost" Then ' (localhost is Microsoft IIS Administration Server Certificate)
                        sData = Reg.ExportKeyToVariable(eHive, "SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates\" & aSubKey(i), False, True, True)
                        AddWarning "New Root certificate is detected! Report to developer, please:" & vbCrLf & Replace(sData, vbCrLf, "\n")
                    End If
                End If
            End If
        Next
    Next
    
    'Set HE = Nothing
    AppendErrorLogCustom "CheckCertificatesEDS - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckCertificatesEDS"
    If inIDE Then Stop: Resume Next
End Sub

Private Sub ParseCertBlob(Blob() As Byte, out_CertHash As String, out_FriendlyName As String, out_IssuedTo As String)
    On Error GoTo ErrorHandler:
    
    'Thanks to Willem Jan Hengeveld
    'https://itsme.home.xs4all.nl/projects/xda/smartphone-certificates.html
    
    Const SHA1_HASH As Long = 3
    Const FRIENDLY_NAME As Long = 11 'Fraudulent
    
    Dim pCertContext    As Long
    Dim CertInfo        As CERT_INFO
    Dim prop            As CERTIFICATE_BLOB_PROPERTY
    
    Dim cStream As clsStream
    Set cStream = New clsStream
    
    out_CertHash = ""
    out_FriendlyName = ""
    out_IssuedTo = ""
    
    'registry blob is an array of CERTIFICATE_BLOB_PROPERTY structures.
    
    cStream.WriteData VarPtr(Blob(0)), UBound(Blob) + 1
    cStream.BufferPointer = 0
    
    Do While cStream.BufferPointer < cStream.Size
        cStream.ReadData VarPtr(prop), 12
        If prop.Length > 0 Then
            ReDim prop.Data(prop.Length - 1)
            cStream.ReadData VarPtr(prop.Data(0)), prop.Length
            
'            Debug.Print "PropID: " & prop.PropertyID
'            Debug.Print "Length: " & prop.length
'            Debug.Print "DataA:   " & Replace(StringFromPtrA(VarPtr(prop.Data(0))), vbNullChar, "-")
'            Debug.Print "DataW:   " & StringFromPtrW(VarPtr(prop.Data(0)))
'            Debug.Print "HexData: " & GetHexStringFromArray(prop.Data)
            
            Select Case prop.PropertyID
            Case SHA1_HASH
                out_CertHash = GetHexStringFromArray(prop.Data)
            Case FRIENDLY_NAME
                out_FriendlyName = StringFromPtrW(VarPtr(prop.Data(0)))
            Case 32
                pCertContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, VarPtr(prop.Data(0)), UBound(prop.Data) + 1)
            
                If pCertContext <> 0 Then
                    
                    If GetCertInfoFromCertificate(pCertContext, CertInfo) Then
                        out_IssuedTo = GetSignerNameFromBLOB(CertInfo.Subject)
                    End If
                    
                    CertFreeCertificateContext pCertContext
                Else
                    Debug.Print "CertCreateCertificateContext failed with 0x" & Hex(Err.LastDllError)
                End If
            End Select
            If out_CertHash <> "" And out_FriendlyName <> "" And out_IssuedTo <> "" Then Exit Do
        End If
    Loop
    
    If out_IssuedTo = "" Then Debug.Print "No SubjectName for cert: " & out_CertHash
    
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

Private Function HexStringToArray(sHexStr As String) As Byte()
    Dim i As Long
    Dim b() As Byte
    
    ReDim b(Len(sHexStr) \ 2 - 1)
    
    For i = 1 To Len(sHexStr) Step 2
        b((i - 1) \ 2) = CLng("&H" & Mid$(sHexStr, i, 2))
    Next
    
    HexStringToArray = b
End Function

Public Sub CheckPolicyACL()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckPolicyACL - Begin"
    
    Dim result As SCAN_RESULT
    Dim i As Long
    Dim SDDL As String, sHit As String
    
    Dim aKey(3) As String
    aKey(0) = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies"
    aKey(1) = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft"
    aKey(2) = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\SystemCertificates"
    aKey(3) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SystemCertificates"
    
    For i = 0 To UBound(aKey)
        If Not CheckKeyAccess(0, aKey(i), KEY_READ) Then
            SDDL = GetKeyStringSD(0, aKey(i))
            
            sHit = "O7 - Policy: Permissions on key are restricted - " & aKey(i) & " - " & SDDL
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    '// TODO: improve it to use correct default permission (no write access for Admin. group / no propagate).
                    AddRegToFix .Reg, RESTORE_KEY_PERMISSIONS_RECURSE, 0, aKey(i)
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    AppendErrorLogCustom "CheckPolicyACL - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckPolicyACL"
    If inIDE Then Stop: Resume Next
End Sub

Sub CheckPolicyScripts()
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckPolicyScripts - Begin"
    
    '
    'For quick overview:
    'gpresult.exe /v - console output of policy scrpits (these doesn't include "Local PC\User" scripts !!! )
    '
    'Keys for analysis:
    '
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Logon\*\*
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Logoff\*\*
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\<SID>\Scripts\Logon\*\*
    'HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\State\<SID>\Scripts\Logon\*\*
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\*\*
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Shutdown\*\*
    'C:\Windows\System32\GroupPolicy\User\Scripts\scripts.ini
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\scripts.ini
    'C:\Windows\System32\GroupPolicy\User\Scripts\psscripts.ini
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\psscripts.ini
    
    'Notice for HKCU:
    '
    'Always represented as:
    ' - Logon\X\Y
    ' - Logoff\X\Y
    '
    'where X defined as:
    '0 - HKCU policy for Local PC (user config)
    '1 - HKCU policy for Local PC\User (user config)*

    ' * How to setup user-specific group policies:
    'https://www.tenforums.com/tutorials/80043-apply-local-group-policy-specific-user-windows-10-a.html
    
    'Notice for HKLM:
    '
    'Always represented as:
    ' - Startup\X\Y
    ' - Shutdown\X\Y
    '
    'where X defined as:
    '0 - HKCU policy for Local PC (machine config)
    '
    'Y - is index number of script record, starting from 0. They are always consecutive*.
    '* non-consecutive indeces brake the chain!
    'When HJT fix the item, it is required to reconstruct the whole chain, so other items become valid.
    
    'Both cases include mirrors in C:\Windows\System32\GroupPolicy location ini-file.
    '* for "Local PC\User" the mirror is located under: C:\Windows\System32\GroupPolicyUsers\<SID>
    
    Dim sHit$, result As SCAN_RESULT
    Dim pos As Long, X As Long, Y As Long, i As Long
    Dim vType As Variant, vFile As Variant
    Dim sFile$, sArgs$, sAlias$, sMD5$, aKeyX$(), aKeyY$(), sKey$, aFiles$(), sIniPath$, sIniPathPS$, sFileSysPath$
    Dim oFiles As clsTrickHashTable
    Set oFiles = New clsTrickHashTable
    oFiles.CompareMode = 1
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    For Each vType In Array("Logon", "Logoff", "Startup", "Shutdown")
        HE.Init HE_HIVE_ALL
        HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\" & vType

        Do While HE.MoveNext

            For X = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aKeyX(), HE.Redirected, False, False, True)

                sFileSysPath = Reg.GetString(HE.Hive, HE.Key & "\" & aKeyX(X), "FileSysPath")
                
                sIniPath = EnvironW(sFileSysPath)
                sIniPath = sFileSysPath & "\Scripts\scripts.ini"
                sIniPathPS = sFileSysPath & "\Scripts\psscripts.ini"
                
                For Y = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key & "\" & aKeyX(X), aKeyY(), HE.Redirected, False, False, True)

                    sKey = HE.Key & "\" & aKeyX(X) & "\" & aKeyY(Y)

                    If Reg.ValueExists(HE.Hive, sKey, "Script", HE.Redirected) Then

                        sFile = Reg.GetString(HE.Hive, sKey, "Script", HE.Redirected)
                        sArgs = Reg.GetString(HE.Hive, sKey, "Parameters", HE.Redirected)

                        If InStr(sFile, ":") = 0 Then 'relative to script storage?
                            sFile = BuildPath(sFileSysPath, "Scripts", vType, sFile)
                        End If

                        sAlias = IIf(bIsWin32, "O7", IIf(HE.Redirected, "O7-32", "O7")) & " - Policy Script: "

                        sFile = FormatFileMissing(sFile)

                        If Not oFiles.Exists(sFile) Then oFiles.Add sFile, 0&

                        sHit = sAlias & HE.HiveNameAndSID & "\..\" & _
                            Replace(sKey, "Software\Microsoft\Windows\CurrentVersion\", "") & _
                            ": [" & "Script" & "] = " & ConcatFileArg(sFile, sArgs)

                        If Not IsOnIgnoreList(sHit) Then

                            If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5

                            With result
                                .Section = "O7"
                                .HitLineW = sHit
                                .Alias = sAlias
                                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, sKey, , , HE.Redirected
                                
                                If 1 = Reg.GetDword(HE.Hive, sKey, "IsPowershell", HE.Redirected) Then
                                    
                                    AddIniToFix .Reg, REMOVE_VALUE_INI, sIniPathPS, vType, aKeyY(Y) & "CmdLine"
                                    AddIniToFix .Reg, REMOVE_VALUE_INI, sIniPathPS, vType, aKeyY(Y) & "Parameters"
                                Else
                                    AddIniToFix .Reg, REMOVE_VALUE_INI, sIniPath, vType, aKeyY(Y) & "CmdLine"
                                    AddIniToFix .Reg, REMOVE_VALUE_INI, sIniPath, vType, aKeyY(Y) & "Parameters"
                                End If
                                
                                'remove state data
                                If vType = "Startup" Or vType = "Shutdown" Then
                                    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Startup\0\0
                                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\" & vType & "\" & aKeyX(X) & "\" & aKeyY(Y), , , HE.Redirected
                                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\" & vType & "\" & aKeyX(X) & "\" & aKeyY(Y), , , HE.Redirected
                                
                                Else ' vType = "Logon" Or vType = "Logoff" Then
                                    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\S-1-5-21-4161311594-4244952198-1204953518-1000\Scripts\Logon\0\0
                                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\" & HE.SID & "\Scripts\" & vType & "\" & aKeyX(X) & "\" & aKeyY(Y), , , HE.Redirected
                                    AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\State\" & HE.SID & "\Scripts\" & vType & "\" & aKeyX(X) & "\" & aKeyY(Y), , , HE.Redirected
                                End If
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                                .CureType = FILE_BASED Or REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If

                Next
            Next
        Loop
    Next
    
    'Checking files inside script folders
    '
    'C:\Windows\System32\GroupPolicy\User\Scripts\Logon
    'C:\Windows\System32\GroupPolicy\User\Scripts\Logoff
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\Startup
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown
    'C:\Windows\System32\GroupPolicyUsers\<SID>\User\Scripts\Logoff
    'C:\Windows\System32\GroupPolicyUsers\<SID>\User\Scripts\Logon
    
    For Each vType In Array("User\Scripts\Logon", "User\Scripts\Logoff", "Machine\Scripts\Startup", "Machine\Scripts\Shutdown")
    
        aFiles = ListFiles(BuildPath(sWinSysDir, "GroupPolicy", vType))
        If AryItems(aFiles) Then
            For i = 0 To UBound(aFiles)
                sFile = aFiles(i)
                If Not oFiles.Exists(sFile) Then 'don't duplicate registry records scan results
                    
                    sAlias = "O7 - Policy Script: "
                    sHit = sAlias & sFile
                    
                    If Not IsOnIgnoreList(sHit) Then
                        
                        If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5
                        
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            .Alias = sAlias
                            AddFileToFix .File, REMOVE_FILE, sFile
                            .CureType = FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        End If
    Next
    
    'C:\Windows\System32\GroupPolicy\User\Scripts\Scripts.ini
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\Scripts.ini
    'C:\Windows\System32\GroupPolicyUsers\<SID>\User\Scripts\Scripts.ini
    'C:\Windows\System32\GroupPolicy\User\Scripts\psscripts.ini
    'C:\Windows\System32\GroupPolicy\Machine\Scripts\psscripts.ini
    'C:\Windows\System32\GroupPolicyUsers\<SID>\User\Scripts\psscripts.ini
    '
    '[Logon]
    '0CmdLine=C:\Users\Alex\Desktop\Alex1.ps1
    '0Parameters=-hi_there
    '1CmdLine=""C:\Users\Alex\Desktop\Alex2.ps1""
    '1Parameters=
    
    Dim cIni As clsIniFile
    Dim nSID As Long, nItem As Long
    Dim sIni As String, vIniName
    Dim aSections(), aSection, aNames(), aName
    
    For Each vIniName In Array("scripts.ini", "psscripts.ini")
    
        For Each vFile In Array( _
            BuildPath(sWinSysDir, "GroupPolicy\User\Scripts\"), _
            BuildPath(sWinSysDir, "GroupPolicy\Machine\Scripts\"), _
            "<SID>")
            
            nSID = LBound(gSID_All)
            Do
                If vFile = "<SID>" Then
                    sIni = BuildPath(sWinSysDir, "GroupPolicyUsers", gSID_All(nSID), "User\Scripts\", vIniName)
                    nSID = nSID + 1
                Else
                    sIni = vFile & vIniName
                End If
                
                If FileExists(sIni) Then
                
                    Set cIni = New clsIniFile
                    cIni.InitFile sIni, 0
                    
                    aSections = cIni.GetSections()
                    
                    For Each aSection In aSections()
                    
                        aNames = cIni.GetParamNames(aSection)
                        
                        For Each aName In aNames()
                        
                            pos = InStr(2, aName, "CmdLine", 1)
                            
                            If pos <> 0 Then
                            
                                If IsNumeric(Left$(aName, pos - 1)) Then
                                
                                    nItem = CLng(Left$(aName, pos - 1))
                                
                                    sFile = cIni.ReadParam(aSection, nItem & "CmdLine")
                                    sArgs = cIni.ReadParam(aSection, nItem & "Parameters")
                                    
                                    sFile = UnQuote(Replace$(sFile, """""", """"))
                                    
                                    If InStr(sFile, ":") = 0 Then 'relative to script storage?
                                        sFile = BuildPath(GetParentDir(sIni), aSection, sFile)
                                    End If
                                    
                                    sFile = FormatFileMissing(sFile)
                                    
                                    If Not oFiles.Exists(sFile) Then 'no such file in registry entries scan
                                        
                                        sAlias = "O7 - Policy Script: "
                                        
                                        sHit = sAlias & sIni & ": " & "[" & aSection & "] " & aName & " = " & ConcatFileArg(sFile, sArgs)
                                        
                                        If Not IsOnIgnoreList(sHit) Then
                                            
                                            If g_bCheckSum Then sMD5 = GetFileCheckSum(sFile): sHit = sHit & sMD5
                                            
                                            With result
                                                .Section = "O7"
                                                .HitLineW = sHit
                                                .Alias = sAlias
                                                AddIniToFix .Reg, REMOVE_VALUE_INI, sIni, aSection, aName
                                                AddIniToFix .Reg, REMOVE_VALUE_INI, sIni, aSection, nItem & "Parameters"
                                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                                                .CureType = FILE_BASED Or INI_BASED
                                            End With
                                            AddToScanResults result
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                    
                End If
                
            Loop Until vFile <> "<SID>" Or nSID > UBound(gSID_All)
        Next
    Next
    
    Set oFiles = Nothing
    
    AppendErrorLogCustom "CheckPolicyScripts - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckPolicyScripts"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub PolicyScripts_RebuildChain()

    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "PolicyScripts_RebuildChain - Begin"
    
    'rebuild registry chain
    
    Dim vType, vFile
    Dim sKey As String, sIniPath As String, sName As String
    Dim aFiles() As String, aKeyX() As String, aKeyY() As String
    Dim X As Long, Y As Long, i As Long, idx As Long
    
    Dim oFiles As clsTrickHashTable
    Set oFiles = New clsTrickHashTable
    oFiles.CompareMode = 1
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    For Each vType In Array("Logon", "Logoff", "Startup", "Shutdown")
        HE.Init HE_HIVE_ALL
        HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\" & vType
        
        Do While HE.MoveNext
        
            For X = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aKeyX(), HE.Redirected, False, False, True)
                
                sIniPath = Reg.GetString(HE.Hive, HE.Key & "\" & aKeyX(X), "FileSysPath") & "\Scripts\scripts.ini"
                sIniPath = EnvironW(sIniPath)
                
                If Not oFiles.Exists(sIniPath) Then oFiles.Add sIniPath, 0&
                
                idx = 0
                
                For Y = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key & "\" & aKeyX(X), aKeyY(), HE.Redirected, False, False, True)
                
                    If aKeyY(Y) <> idx Then
                        
                        sKey = HE.Key & "\" & aKeyX(X) & "\" & aKeyY(Y)
                        
                        Reg.RenameKey HE.Hive, sKey, CStr(idx), HE.Redirected
                    End If
                    
                    idx = idx + 1
                Next
            Next
        Loop
    Next
    
    'rebuild ini chain
    
    Dim cIni As clsIniFile
    
    'preparing list of ini files
    aFiles = ListFiles(BuildPath(sWinSysDir, "GroupPolicyUsers"), ".ini", True)
    
    If AryItems(aFiles) Then
        ReDim Preserve aFiles(UBound(aFiles) + 4)
    Else
        ReDim aFiles(3)
    End If
    aFiles(UBound(aFiles) - 3) = BuildPath(sWinSysDir, "GroupPolicy\User\Scripts\psscripts.ini")
    aFiles(UBound(aFiles) - 2) = BuildPath(sWinSysDir, "GroupPolicy\User\Scripts\scripts.ini")
    aFiles(UBound(aFiles) - 1) = BuildPath(sWinSysDir, "GroupPolicy\Machine\Scripts\psscripts.ini")
    aFiles(UBound(aFiles) - 0) = BuildPath(sWinSysDir, "GroupPolicy\Machine\Scripts\scripts.ini")
    For i = 0 To UBound(aFiles)
    
        sName = GetFileName(aFiles(i), True)
        
        If StrComp(sName, "scripts.ini", 1) = 0 _
            Or StrComp(sName, "psscripts.ini", 1) = 0 Then
        
            'append to files pointed by registry entries
            If Not oFiles.Exists(aFiles(i)) Then oFiles.Add aFiles(i), 0&
        End If
    Next
    
    Dim aParam() As Variant
    Dim iNum As Long
    
    For Each vFile In oFiles.Keys
        
        If FileExists(vFile) Then
        
            Set cIni = New clsIniFile
            cIni.InitFile CStr(vFile), 0
            
            'Debug.Print vFile
            
            For Each vType In Array("Logon", "Logoff", "Startup", "Shutdown")
                
                idx = -1
                
                aParam = cIni.GetParamNames(vType)
                
                For X = 0 To UBoundSafe(aParam)
                    
                    sName = Mid$(aParam(X), 2)
                    
                    If StrComp(sName, "CmdLine", 1) = 0 Then
                        
                        If IsNumeric(Left$(aParam(X), 1)) Then
                            
                            idx = idx + 1
                            iNum = CLng(Left$(aParam(X), 1))
                            
                            If iNum <> idx Then
                                cIni.RenameParam vType, CStr(iNum) & "CmdLine", CStr(idx) & "CmdLine"
                                cIni.RenameParam vType, CStr(iNum) & "Parameters", CStr(idx) & "Parameters"
                            End If
                        End If
                    End If
                    
                    'Debug.Print aParam(X) & "  =  " & cIni.ReadParam(vType, aParam(X))
                Next
            Next
            
            Set cIni = Nothing
            
        End If
    Next
    
    Set oFiles = Nothing
    
    AppendErrorLogCustom "PolicyScripts_RebuildChain - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "PolicyScripts_RebuildChain"
    If inIDE Then Stop: Resume Next
End Sub


Public Sub CheckPolicies()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckPolicies - Begin"
    
    '//TODO:
    'HKEY_CURRENT_USER\Software\Policies\Microsoft
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies

    'UAC (EnableLUA, PromptOnSecureDesktop ...) - http://www.oszone.net/11424
    
    Dim sDrv As String, aValue() As String, i&, lData&, bData() As Byte
    Dim sHit$, result As SCAN_RESULT
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    
    aValue = Split("DisableRegistryTools|DisableTaskMgr", "|")
    
    Do While HE.MoveNext
        'key - x64 Shared
        For i = 0 To UBound(aValue)
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i))
            If lData <> 0 Then
                sHit = "O7 - Policy: " & HE.HiveNameAndSID & "\..\Policies\System: " & "[" & aValue(i) & "] = " & lData
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    'Taskbar policies
    'см. Клименко Р. Тонкости реестра Windows Vista. Трюки и эффекты.
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    aValue = Split("NoSetTaskbar|TaskbarLockAll|LockTaskbar|NoTrayItemsDisplay|NoChangeStartMenu|NoStartMenuMorePrograms|NoRun" & _
        "NoSMConfigurePrograms|NoViewOnDrive|RestrictRun|DisallowRun|NoControlPanel|NoDispCpl", "|")
    
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i))
            If lData <> 0 Then
                sHit = "O7 - Taskbar policy: " & HE.HiveNameAndSID & "\..\Policies\Explorer: [" & aValue(i) & "] = " & lData
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O7"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    'hide drives in My PC window
    'https://support.microsoft.com/en-us/help/555438
    HE.Repeat
    Do While HE.MoveNext
        If Reg.ValueExists(HE.Hive, HE.Key, "NoDrives") Then
            bData = Reg.GetBinary(HE.Hive, HE.Key, "NoDrives")
            If AryItems(bData) Then
                GetMem4 bData(0), lData
                For i = 65 To 90
                    If lData And (2 ^ (i - 65)) Then
                        sDrv = sDrv & Chr$(i) & ", "
                    End If
                Next
                If Len(sDrv) <> 0 Then
                    sDrv = Left$(sDrv, Len(sDrv) - 2)
                    sHit = "O7 - Explorer Policy: " & HE.HiveNameAndSID & "\..\Policies\Explorer: [NoDrives] = 0x" & ByteArrayToHex(bData) & _
                        " (Disk: " & sDrv & ")"
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            .Reboot = True
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "NoDrives"
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
        End If
    Loop
    
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    
    aValue = Split("Start_ShowRun", "|")
    
    Do While HE.MoveNext
        For i = 0 To UBound(aValue)
            If Reg.ValueExists(HE.Hive, HE.Key, aValue(i)) Then
                lData = Reg.GetDword(HE.Hive, HE.Key, aValue(i))
                If lData = 0 Then
                    sHit = "O7 - Taskbar policy: " & HE.HiveNameAndSID & "\..\Explorer\Advanced: [" & aValue(i) & "] = " & lData
                
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O7"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            End If
        Next
    Loop
    
    'https://blog.malwarebytes.com/detections/pum-optional-disallowrun/
    '"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    Dim iEnabled As Long
    Dim sData As String
    
    HE.Init HE_HIVE_ALL, , HE_REDIR_NO_WOW
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun"
    
    Do While HE.MoveNext
    
        iEnabled = Reg.GetDword(HE.Hive, HE.Key, "DisallowRun", HE.Redirected)
        
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue(), HE.Redirected)
   
            sData = Reg.GetData(HE.Hive, HE.Key, aValue(i), HE.Redirected)
            
            sHit = "O7 - Policy: " & HE.HiveNameAndSID & "\..\Policies\Explorer\DisallowRun: " & "[" & aValue(i) & "] = " & sData & IIf(iEnabled = 0, " (disabled)", "")
            
            If Not IsOnIgnoreList(sHit) Then
                With result
                    .Section = "O7"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, aValue(i)
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        Next
    Loop

    AppendErrorLogCustom "CheckPolicies - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckPolicies"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO7Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO7Item - Begin"

    'Policies
    CheckPolicies
    
    'Policy - Logon scripts
    CheckPolicyScripts
    
    'Untrusted certificates
    UpdateProgressBar "O7-Cert"
    Call CheckCertificatesEDS
    If Not bAutoLogSilent Then DoEvents
    
    ' System troubleshooting
    UpdateProgressBar "O7-Trouble"
    Call CheckSystemProblems '%temp%, %tmp%, disk free space < 1 GB.
    If Not bAutoLogSilent Then DoEvents
    
    'Check for DACL lock on Policy key
    UpdateProgressBar "O7-ACL"
    Call CheckPolicyACL
    If Not bAutoLogSilent Then DoEvents
    
    'IP Security
    UpdateProgressBar "O7-IPSec"
    Call CheckIPSec
    
    AppendErrorLogCustom "CheckO7Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO7Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckIPSec()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckIPSec - Begin"
    
    Dim sHit$, sHit1$, sHit2$, result As SCAN_RESULT
    Dim i As Long
    
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
    Dim KeyISAKMP As String, j As Long, KeyFilter() As String, k As Long, NegAction As String, NegType As String, bEnabled As Boolean, sActPolicy As String
    Dim bRegexpInit As Boolean, bFilterData() As Byte, IP(1) As String, RuleAction As String, bMirror As Boolean, DataSerialized As String
    Dim Packet_Type(1) As String, M As Long, n As Long, PortNum(1) As Long, ProtocolType As String, idxBaseOffset As Long, IpFil As IPSEC_FILTER_RECORD, RecCnt As Byte
    Dim oMatches As IRegExpMatchCollection, IPTypeFlag(1) As Long, b() As Byte, bAtLeastOneFilter As Boolean, bNoFilter As Boolean
    Dim bSafe As Boolean
    Dim odHit As clsTrickHashTable
    Set odHit = New clsTrickHashTable
    
    
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
        
        If AryItems(KeyNFA) Then
            
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
                        
            If 0 = AryItems(KeyFilter) Then
            
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
                
                For k = 0 To UBound(KeyFilter)
                    KeyFilter(k) = MidFromCharRev(KeyFilter(k), "\")
                    KeyFilter(k) = IIf(KeyFilter(k) = "", "", "HKLM\SOFTWARE\Policies\Microsoft\Windows\IPSec\Policy\Local\" & KeyFilter(k))
                Next
                
                For k = 0 To UBound(KeyFilter)
                    
                    Erase IP
                    Erase Packet_Type: Packet_Type(0) = "Unknown": Packet_Type(1) = "Unknown"
                    Erase PortNum
                    ProtocolType = ""
                    bMirror = False
                    
                    bFilterData() = Reg.GetBinary(0&, KeyFilter(k), "ipsecData")
                    
                    If AryItems(bFilterData) Then

                      AppendErrorLogCustom "CheckO7Item - Regexp - Begin"

                      If Not g_bRegexpInit Then
                        Set oRegexp = New cRegExp
                        g_bRegexpInit = True
                      End If

                      If Not bRegexpInit Then
                        bRegexpInit = True
                        oRegexp.IgnoreCase = True
                        oRegexp.Global = True
                        oRegexp.Pattern = "(00|01)(000000)(........)(00000000|FFFFFFFF)(........)(00000000|FFFFFFFF)(00000000)(((06|11)000000........)|((00|01|06|08|11|14|16|1B|42|FF|..)00000000000000))00(00|01|02|03|04|81|82|83|84)0000"
                      End If
                      
                      Set oMatches = oRegexp.Execute(SerializeByteArray(bFilterData))
                    
                      AppendErrorLogCustom "CheckO7Item - Regexp - End"
    
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
    
    Set odHit = Nothing
    
    If Not bAutoLogSilent Then DoEvents
    
    AppendErrorLogCustom "CheckIPSec - End"
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
    
    'do not show in log disabled entries
    If Not bIgnoreAllWhitelists Then
        If Not bEnabled Then Return
    End If
    
    sHit1 = "O7 - IPSec: Name: " & IPSecName & " " & _
        "(" & Format$(dModify, "yyyy\/mm\/dd") & ")" & " - "
        
    sHit2 = IPSecID & " - " & _
        IIf(bNoFilter, "No rules ", _
        "Source: " & IIf(Packet_Type(0) = "IP", "IP: " & IP(0), Packet_Type(0)) & _
        IIf((ProtocolType = "TCP" Or ProtocolType = "UDP") And PortNum(0) <> 0, " (Port " & PortNum(0) & " " & ProtocolType & ")", "") & " - " & _
        "Destination: " & IIf(Packet_Type(1) = "IP", "IP: " & IP(1), Packet_Type(1)) & _
        IIf((ProtocolType = "TCP" Or ProtocolType = "UDP") And PortNum(1) <> 0, " (Port " & PortNum(1) & " " & ProtocolType & ")", "") & " " & _
        IIf(bMirror, "(mirrored) ", "")) & "- Action: " & RuleAction & IIf(bEnabled, "", " (disabled)")
    
    sHit = sHit1 & sHit2
    
    If odHit.Exists(sHit) Then 'skip several identical rules
        If Not bAutoLogSilent Then DoEvents
        Return
    Else
        odHit.Add sHit, 0&
    End If
    
    bSafe = False
    
    'Whitelists
    If bHideMicrosoft Then
      If (OSver.MajorMinor <= 5.2) Or (OSver.MajorMinor = 5.2 And OSver.IsWin64) Then  'Win2k / XP / XP x64
        If StrComp(sHit2, "{72385236-70fa-11d1-864c-14a300000000} - No rules - Action: Default response (disabled)", 1) = 0 Then
            bSafe = True
        ElseIf StrComp(sHit2, "{72385230-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Allow (disabled)", 1) = 0 Then
            bSafe = True
        ElseIf StrComp(sHit2, "{72385230-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Inbound pass-through (disabled)", 1) = 0 Then
            bSafe = True
        ElseIf StrComp(sHit2, "{7238523c-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Allow (disabled)", 1) = 0 Then
            bSafe = True
        ElseIf StrComp(sHit2, "{7238523c-70fa-11d1-864c-14a300000000} - Source: my IP - Destination: Any IP (mirrored) - Action: Inbound pass-through (disabled)", 1) = 0 Then
            bSafe = True
        End If
      End If
    End If
    
    If Not bSafe Then
      If Not IsOnIgnoreList(sHit) Then
        With result
            .Section = "O7"
            .HitLineW = sHit
            AddRegToFix .Reg, REMOVE_KEY, 0, KeyPolicy(i)
            If KeyISAKMP <> "" Then AddRegToFix .Reg, REMOVE_KEY, 0, KeyISAKMP
            If AryItems(KeyNFA) Then
                For M = 0 To UBound(KeyNFA)
                    If KeyNFA(M) <> "" Then
                        AddRegToFix .Reg, REMOVE_KEY, 0, KeyNFA(M)
                    End If
                    If KeyNegotiation(M) <> "" Then
                        AddRegToFix .Reg, REMOVE_KEY, 0, KeyNegotiation(M)
                    End If
                Next
            End If
            If AryItems(KeyFilter) Then
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
        AddToScanResults result
      End If
    End If
    
    Return
ErrorHandler:
    ErrorMsg Err, "ModMain_CheckIPSec"
    If inIDE Then Stop: Resume Next
End Sub

'byte array -> to Hex String
Public Function SerializeByteArray(b() As Byte, Optional Delimiter As String = "") As String
    Dim i As Long
    Dim s As String
    SerializeByteArray = String$((UBound(b) + 1) * 2, "0")
       
    For i = 0 To UBound(b)
        s = Hex$(b(i))
        Mid$(SerializeByteArray, (i * 2) + 1 + IIf(Len(s) = 2, 0, 1)) = s
    Next
End Function

'Serialized Hex String of bytes -> byte array
Public Function DeSerializeToByteArray(s As String, Optional Delimiter As String = "") As Byte()
    On Error GoTo ErrorHandler:
    Dim i As Long
    Dim n As Long
    Dim b() As Byte
    Dim ArSize As Long
    If Len(s) = 0 Then Exit Function
    ArSize = (Len(s) + Len(Delimiter)) \ (2 + Len(Delimiter)) '2 chars on byte + add final delimiter
    ReDim b(ArSize - 1) As Byte
    For i = 1 To Len(s) Step 2 + Len(Delimiter)
        b(n) = CLng("&H" & Mid$(s, i, 2))
        n = n + 1
    Next
    DeSerializeToByteArray = b
    Exit Function
ErrorHandler:
    Debug.Print "Error in DeSerializeByteString"
End Function

Public Sub FixO7Item(sItem$, result As SCAN_RESULT)
    'O7 - Disabling of Policies
    On Error GoTo ErrorHandler:
    
    If InStr(1, result.HitLineW, "Policy Script:", 1) <> 0 Then bNeedRebuildPolicyChain = True
    
    If result.CureType = CUSTOM_BASED Then
    
        If InStr(1, result.HitLineW, "Free disk space", 1) <> 0 Then
            RunCleanMgr
            
        ElseIf InStr(1, result.HitLineW, "Computer name (hostname) is not set", 1) <> 0 Then
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
        FixRegistryHandler result
        FixFileHandler result
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
    
    '//TODO: (от sov44)
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
    
    Dim hKey&, i&, sName$, lpcName&, sFile$, sHit$, result As SCAN_RESULT, pos&, bSafe As Boolean
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
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
        
                sHit = "O8 - Context menu item: " & HE.HiveNameAndSID & "\..\Internet Explorer\MenuExt\" & sName & ": (default) = " & sFile
                
                bSafe = False
                If WhiteListed(sFile, "EXCEL.EXE", True) Then bSafe = True 'MS Office
                If WhiteListed(sFile, "ONBttnIE.dll", True) Then bSafe = True 'MS Office
                
                If Not IsOnIgnoreList(sHit) And (Not bSafe) Then
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    With result
                        .Section = "O8"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
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

Public Sub FixO8Item(sItem$, result As SCAN_RESULT)
    'O8 - Extra context menu items
    'O8 - Extra context menu item: [name] - html file
    'HKCU\Software\Microsoft\Internet Explorer\MenuExt
    
    On Error GoTo ErrorHandler:
    
    FixRegistryHandler result
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
    
    Dim hKey&, i&, sData$, sCLSID$, sCLSID2$, lpcName&, sFile$, sHit$, sBuf$, result As SCAN_RESULT
    Dim pos&, bSafe As Boolean
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
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
                If InStr(1, sFile, "res://", vbTextCompare) = 1 Then
                    'And (LCase$(Right$(sFile, 4)) = ".htm" Or LCase$(Right$(sFile, 4)) = "html") Then
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
                " - Button: " & HE.HiveNameAndSID & "\..\" & sCLSID & ": " & sData & " - " & sFile
              
              If Not IsOnIgnoreList(sHit) Then
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                With result
                    .Section = "O9"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sCLSID, , , HE.Redirected
                    AddRegToFix .Reg, REMOVE_VALUE, HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", sCLSID
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
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
                  " - Tools menu item: " & HE.HiveNameAndSID & "\..\" & sCLSID & ": " & sData & " - " & sFile
                If Not IsOnIgnoreList(sHit) Then
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    With result
                        .Section = "O9"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sCLSID, , , HE.Redirected
                        AddRegToFix .Reg, REMOVE_VALUE, HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\LowRegistry\Extensions\CmdMapping", sCLSID
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
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

Public Sub FixO9Item(sItem$, result As SCAN_RESULT)
    'O9 - Extra buttons/Tools menu items
    'O9 - Extra button: [name] - [CLSID] - [file] [(HKCU)]
    
    On Error GoTo ErrorHandler:

    FixRegistryHandler result
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
    Dim hKey&, i&, sSubKey$, sName$, lpcName&, sHit$, result As SCAN_RESULT
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
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
                  "TOEGANKELIJKHEID.TABS.INTERNATIONAL*.ACCELERATED_GRAPHICS", sSubKey) = 0 Or bIgnoreAllWhitelists Then
                  
                    sName = Reg.GetString(HE.Hive, HE.Key & "\" & sSubKey, "Text", HE.Redirected)
                  
                    If Len(sName) <> 0 Then
                        'O11 - Options group:
                        'O11-32 - Options group:
                        sHit = IIf(bIsWin32, "O11", IIf(HE.Redirected, "O11-32", "O11")) & _
                          " - " & HE.HiveNameAndSID & "\..\Internet Explorer\AdvancedOptions\" & sSubKey & ": [Text] = " & sName
                
                        If bIgnoreAllWhitelists Or Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O11"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sSubKey, , , HE.Redirected
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
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

Public Sub FixO11Item(sItem$, result As SCAN_RESULT)
    'O11 - Options group: [BLA] Blah"
    On Error GoTo ErrorHandler:
    FixRegistryHandler result
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
    
    Dim hKey&, i&, sName$, sData$, sFile$, sArgs$, sHit$, lpcName&, result As SCAN_RESULT
    Dim aKey() As String, aDes() As String
    ReDim aKey(1), aDes(1)
    
    aKey(0) = "Software\Microsoft\Internet Explorer\Plugins\Extension"
    aDes(0) = "Internet Explorer\Plugins\Extension"
    
    aKey(1) = "Software\Microsoft\Internet Explorer\Plugins\MIME"
    aDes(1) = "Internet Explorer\Plugins\MIME"
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKeys aKey
    
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
              HE.HiveNameAndSID & "\..\" & aDes(HE.KeyIndex) & "\" & sName & ": [Location] = " & ConcatFileArg(sFile, sArgs)
              
            If Not IsOnIgnoreList(sHit) Then
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                With result
                    .Section = "O12"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName, , , HE.Redirected
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                    .CureType = REGISTRY_BASED Or FILE_BASED
                End With
                AddToScanResults result
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

Public Sub FixO12Item(sItem$, result As SCAN_RESULT)
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
    
    FixRegistryHandler result
    FixFileHandler result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO12Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO13Item()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO13Item - Begin"
    
    Dim sDummy$, sHit$, result As SCAN_RESULT
    Dim aKey() As String, aVal() As String, aExa() As String, aDes() As String, i As Long
    
    ReDim aKey(6)
    ReDim aVal(UBound(aKey))
    ReDim aExa(UBound(aKey))
    ReDim aDes(UBound(aKey))
    
    aKey(0) = "DefaultPrefix"
    aVal(0) = ""
    aExa(0) = "http://"
    'aDes(0) = "DefaultPrefix"
    
    aKey(1) = "Prefixes"
    aVal(1) = "www"
    aExa(1) = "http://"
    'aDes(1) = "WWW Prefix"
    
    aKey(2) = "Prefixes"
    aVal(2) = "www."
    aExa(2) = ""
    'aDes(2) = "WWW. Prefix"
    
    aKey(3) = "Prefixes"
    aVal(3) = "home"
    aExa(3) = "http://"
    'aDes(3) = "Home Prefix"
    
    aKey(4) = "Prefixes"
    aVal(4) = "mosaic"
    aExa(4) = "http://"
    'aDes(4) = "Mosaic Prefix"
    
    aKey(5) = "Prefixes"
    aVal(5) = "ftp"
    aExa(5) = "ftp://"
    'aDes(5) = "FTP Prefix"
    
    aKey(6) = "Prefixes"
    aVal(6) = "gopher"
    aExa(6) = "gopher://|"
    'aDes(6) = "Gopher Prefix"
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\URL"

    Do While HE.MoveNext
    
        For i = 0 To UBound(aKey)
        
            sDummy = Reg.GetString(HE.Hive, HE.Key & "\" & aKey(i), aVal(i), HE.Redirected)
            
            'exclude empty HKCU / HKU
            If Not (HE.Hive <> HKLM And sDummy = "") Then
            
                If Not inArraySerialized(sDummy, aExa(i), "|", , , vbBinaryCompare) Or Not bHideMicrosoft Then
                    
                    sHit = IIf(bIsWin32, "O13", IIf(HE.Redirected, "O13-32", "O13")) & " - " & HE.HiveNameAndSID & "\..\URL\" & aKey(i) & _
                        ": [" & aVal(i) & "] = " & sDummy
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O13"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key & "\" & aKey(i), aVal(i), SplitSafe(aExa(i), "|")(0), HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
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

Public Sub FixO13Item(sItem$, result As SCAN_RESULT)
    'defaultprefix fix
    'O13 - DefaultPrefix: http://www.hijacker.com/redir.cgi?
    'O13 - [WWW/Home/Mosaic/FTP/Gopher] Prefix: ..
    
    On Error GoTo ErrorHandler:
    FixRegistryHandler result
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
    
    aLogStrings = ReadFileToArray(sFile, FileGetTypeBOM(sFile) = 1200)
    
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
    If (sSearchAssis <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm" And _
      sSearchAssis <> g_DEFSEARCHASS And sSearchAssis <> "") Or Not bHideMicrosoft Then
        sHit = "O14 - IERESET.INF: SearchAssistant = " & sSearchAssis
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'CustomizeSearch = http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm
    If (sCustSearch <> "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm" And _
      sCustSearch <> g_DEFSEARCHCUST And sCustSearch <> "") Or Not bHideMicrosoft Then
        sHit = "O14 - IERESET.INF: CustomizeSearch = " & sCustSearch
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'SEARCH_PAGE_URL = http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch
    If (sSearchPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&ar=iesearch" And _
      sSearchPage <> "http://www.msn.com" And _
      sSearchPage <> "https://www.msn.com" And _
      sSearchPage <> g_DEFSEARCHPAGE) Or Not bHideMicrosoft Then
        sHit = "O14 - IERESET.INF: [Strings] SEARCH_PAGE_URL = " & sSearchPage
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'START_PAGE_URL  = http://www.msn.com
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '                  http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If (sStartPage <> "http://www.msn.com" And _
       sStartPage <> "https://www.msn.com" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
       sStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome" And _
       sStartPage <> g_DEFSTARTPAGE) Or Not bHideMicrosoft Then
        sHit = "O14 - IERESET.INF: [Strings] START_PAGE_URL = " & sStartPage
        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
    End If
    
    'MS_START_PAGE_URL=http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome
    '(=START_PAGE_URL) http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome
    If sMsStartPage <> vbNullString Then
        If (sMsStartPage <> "http://www.msn.com" And _
           sMsStartPage <> "https://www.msn.com" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.5&ar=msnhome" And _
           sMsStartPage <> "http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=6&ar=msnhome" And _
           sMsStartPage <> g_DEFSTARTPAGE) Or Not bHideMicrosoft Then
            sHit = "O14 - IERESET.INF: [Strings] MS_START_PAGE_URL = " & sMsStartPage
            If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O14", sHit
        End If
    End If
    
    AppendErrorLogCustom "CheckO14Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO14Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO14Item(sItem$, result As SCAN_RESULT)
    'resetwebsettings fix
    'O14 - IERESET.INF: [item]=[URL]
    
    On Error GoTo ErrorHandler:
    'sItem - not used
    Dim sLine$, sFixedIeResetInf$, ff%
    Dim i&, aLogStrings() As String, sFile$, isUnicode As Boolean
    
    sFile = sWinDir & "\INF\iereset.inf"
    
    If Not FileExists(sFile) Then Exit Sub
    
    BackupFile result, sFile
    
    isUnicode = (FileGetTypeBOM(sFile) = 1200)
    aLogStrings = ReadFileToArray(sFile, IIf(isUnicode, True, False))
    
    For i = 0 To UBound(aLogStrings)
        sLine = aLogStrings(i)

            If InStr(sLine, "SearchAssistant") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""SearchAssistant"",0,""" & _
                    IIf(g_DEFSEARCHASS <> "", g_DEFSEARCHASS, "") & """" & vbCrLf
            ElseIf InStr(sLine, "CustomizeSearch") > 0 Then
                sFixedIeResetInf = sFixedIeResetInf & "HKLM,""Software\Microsoft\Internet Explorer\Search"",""CustomizeSearch"",0,""" & _
                    IIf(g_DEFSEARCHCUST <> "", g_DEFSEARCHCUST, "") & """" & vbCrLf
            ElseIf InStr(sLine, "START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "START_PAGE_URL=""" & _
                    IIf(g_DEFSTARTPAGE <> "", g_DEFSTARTPAGE, "https://www.msn.com") & """" & vbCrLf
            ElseIf InStr(sLine, "SEARCH_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "SEARCH_PAGE_URL=""" & _
                    IIf(g_DEFSEARCHPAGE <> "", g_DEFSEARCHPAGE, "https://www.msn.com") & """" & vbCrLf
            ElseIf InStr(sLine, "MS_START_PAGE_URL=") = 1 Then
                sFixedIeResetInf = sFixedIeResetInf & "MS_START_PAGE_URL=""" & _
                    IIf(g_DEFSTARTPAGE <> "", g_DEFSTARTPAGE, "https://www.msn.com") & """" & vbCrLf
            Else
                sFixedIeResetInf = sFixedIeResetInf & sLine & vbCrLf
            End If
        
    Next
    sFixedIeResetInf = Left$(sFixedIeResetInf, Len(sFixedIeResetInf) - 2)   '-CrLf
    
    DeleteFileWEx (StrPtr(sFile))
    
    ff = FreeFile()
    
    If isUnicode Then
        Dim b() As Byte
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
    Dim i&, j&, sHit$, sAlias$, sIPRange$, bSafe As Boolean, result As SCAN_RESULT
    Dim dURL As clsTrickHashTable, aResult() As SCAN_RESULT, iRes As Long, iCur As Long, sURL As String
    Set dURL = New clsTrickHashTable
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
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
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - " & HE.HiveNameAndSID & "\..\ESC Trusted Zone: "
            Else
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - " & HE.HiveNameAndSID & "\..\Trusted Zone: "
            End If
            For i = 0 To UBound(sDomains)
                bSafe = False
                If bHideMicrosoft And Not bIgnoreAllWhitelists Then
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
                                
                                    bSafe = False
                                    If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                                        bSafe = StrBeginWithArray(sProtPrefix & sSubDomains(j) & "." & sDomains(i), aSafeRegDomains)
                                    End If
                                    
                                    If Not bSafe Then
                                        sURL = sSubDomains(j) & "." & sDomains(i)
                                        sAlias = "O15 - Trusted Zone: "
                                        sHit = sAlias & sProtPrefix & sURL
                                        
                                        If Not IsOnIgnoreList(sHit) Then
                                        
                                            'concat several identical URLs to single log line
                                            If dURL.Exists(sURL) Then
                                                iCur = dURL(sURL)
                                            Else
                                                iRes = iRes + 1
                                                iCur = iRes
                                                ReDim Preserve aResult(iRes)
                                                dURL.Add sURL, iCur
                                            End If
                                            
                                            With aResult(iCur)
                                                .Section = "O15"
                                                .HitLineW = sHit
                                                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i) & "\" & sSubDomains(j), , , HE.Redirected
                                                .CureType = REGISTRY_BASED
                                            End With
                                            'AddToScanResults result
                                        End If
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
                        
                            bSafe = False
                            If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                                If StrBeginWithArray(sProtPrefix & sDomains(i), aSafeRegDomains) Then bSafe = True
                            End If
                        
                            If Not bSafe Then
                                sURL = sDomains(i)
                                sAlias = "O15 - Trusted Zone: "
                                sHit = sAlias & sProtPrefix & sURL
                                
                                If Not IsOnIgnoreList(sHit) Then
                                
                                    'concat several identical URLs to single log line
                                    If dURL.Exists(sURL) Then
                                        iCur = dURL(sURL)
                                    Else
                                        iRes = iRes + 1
                                        iCur = iRes
                                        ReDim Preserve aResult(iRes)
                                        dURL.Add sURL, iCur
                                    End If
                                
                                    With aResult(iCur)
                                        .Section = "O15"
                                        .HitLineW = sHit
                                        AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i), , , HE.Redirected
                                        .CureType = REGISTRY_BASED
                                    End With
                                    'AddToScanResults result
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
    Loop
    
    For i = 1 To iRes
        AddToScanResults aResult(i)
    Next
        
    Set dURL = Nothing
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscRanges"
    
    'enum all IP ranges
    Do While HE.MoveNext
        sDomains = Split(Reg.EnumSubKeys(HE.Hive, HE.Key, HE.Redirected), "|")
        If UBound(sDomains) > -1 Then
            If StrEndWith(HE.Key, "EscRanges") Then
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - " & HE.HiveNameAndSID & "\..\ESC Trusted IP range: "
            Else
                sAlias = IIf(HE.Redirected, "O15-32", "O15") & " - " & HE.HiveNameAndSID & "\..\Trusted IP range: "
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
                            sHit = sAlias & sProtPrefix & sIPRange
                            If Not IsOnIgnoreList(sHit) Then
                                With result
                                    .Section = "O15"
                                    .HitLineW = sHit
                                    AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sDomains(i), , , HE.Redirected
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults result
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
    
    '0 = My Computer
    '1 = Intranet
    '2 = Trusted
    '3 = Internet
    '4 = Restricted
    '5 = Unknown
    
    'binding protocol to zone
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
        HE.Init HE_HIVE_ALL, HE_SID_USER Or HE_SID_NO_VIRTUAL
    ElseIf OSver.MajorMinor = 5.2 And OSver.IsServer Then 'Win 2003, 2003 R2
        HE.Init HE_HIVE_ALL, HE_SID_USER Or HE_SID_NO_VIRTUAL
    Else 'Vista+
        HE.Init HE_HIVE_ALL, HE_SID_DEFAULT Or HE_SID_USER Or HE_SID_NO_VIRTUAL
    End If
    
    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\ProtocolDefaults"
    
    If OSver.IsWindows7OrGreater Then
        LastIndex = 11
    ElseIf OSver.IsWindowsVistaOrGreater Then
        LastIndex = 10
    ElseIf OSver.IsWindowsXPOrGreater Then
        LastIndex = 5
    Else
        LastIndex = 4
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
                    If HE.UserName = "UpdatusUser" Then bSafe = True
                    If HE.UserName = "unknown" Then bSafe = True
                End If
            End If
            
            If Not bSafe Then
                If lProtZones(i) < 0 Or lProtZones(i) > 5 Then lProtZones(i) = 5 'Unknown
                If lProtZones(i) = 5 Then
                    If InStr(1, HE.UserName, "MSSQL", 1) <> 0 Then
                        bSafe = True
                    ElseIf InStr(1, HE.UserName, "MsDtsServer", 1) <> 0 Then
                        bSafe = True
                    ElseIf InStr(1, HE.UserName, "defaultuser0", 1) <> 0 Then
                        bSafe = True
                    ElseIf InStr(1, HE.UserName, "Acronis Agent User", 1) <> 0 Then
                        bSafe = True
                        '// TODO: improve it (logon as service)
                    ElseIf Reg.GetDword(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & HE.SID, "State") = 0 Then
                        bSafe = True
                    End If
                End If
                
                If Not bSafe And (lProtZones(i) <> lProtZoneDefs(i)) Then 'check for legit
                    
                    sHit = IIf(HE.Redirected, "O15-32", "O15") & " - " & HE.HiveNameAndSID & "\..\ProtocolDefaults: " & _
                        " - [" & sProtVals(i) & "] protocol is in " & sZoneNames(lProtZones(i)) & " Zone, should be " & sZoneNames(lProtZoneDefs(i)) & " Zone" & _
                        IIf(HE.IsSidUser, " (User: '" & HE.UserName & "')", "")
                        
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O15"
                            .HitLineW = sHit
                            AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, sProtVals(i), lProtZoneDefs(i), HE.Redirected
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
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

Public Sub FixO15Item(sItem$, result As SCAN_RESULT)
'    'O15 - Trusted Zone: free.aol.com (HKLM)
'    'O15 - Trusted Zone: http://free.aol.com
'    'O15 - Trusted IP range: 66.66.66.66 (HKLM)
'    'O15 - Trusted IP range: http://66.66.66.*
'    'O15 - ESC Trusted Zone: free.aol.com (HKLM)
'    'O15 - ESC Trusted IP range: 66.66.66.66
'    'O15 - ProtocolDefaults: 'http' protocol is in Trusted Zone, should be Internet Zone (HKLM)

    On Error GoTo ErrorHandler:
    FixRegistryHandler result
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
    
    Dim sName$, sFriendlyName$, sCodebase$, i&, j&, hKey&, lpcName&, sHit$, result As SCAN_RESULT
    Dim sOSD$, sInf$, sInProcServer32$, aValue() As String
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
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
                  Or Not bHideMicrosoft Then
           
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
                        sCodebase = FormatFileMissing(PathNormalize(sCodebase))
                    End If
                    
                    ' "O16 - DPF: "
                    ' CODEBASE - is a URL
                    sHit = IIf(bIsWin32, "O16", IIf(HE.Redirected, "O16-32", "O16")) & " - DPF: " & HE.HiveNameAndSID & "\..\" & _
                      sName & "\DownloadInformation: " & sFriendlyName & " [CODEBASE] = " & sCodebase
                    
                    If g_bCheckSum Then
                        'if file
                        If Mid$(sCodebase, 2, 1) = ":" Then sHit = sHit & GetFileCheckSum(sCodebase)
                    End If
                    
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O16"
                            .HitLineW = sHit
                            
                            sOSD = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\DownloadInformation", "OSD", HE.Redirected)
                            sInf = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\DownloadInformation", "INF", HE.Redirected)

                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sInProcServer32 'Or UNREG_DLL
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sOSD
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sInf
                            
                            AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sName, , , REG_REDIRECTION_BOTH
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & sName, , , HE.Redirected
                            
                            AddRegToFix .Reg, REMOVE_KEY, HKLM, "SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility\" & sName, , , HE.Redirected
                            AddRegToFix .Reg, REMOVE_KEY, HKCU, "Software\Microsoft\Windows\CurrentVersion\Ext\Stats\" & sName, , , HE.Redirected
                            
                            For j = 1 To Reg.EnumValuesToArray(HKLM, HE.Key & "\" & sName & "\Contains\Files", aValue, HE.Redirected)
                                AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, aValue(j)
                            Next
                            
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
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

Public Sub FixO16Item(sItem$, result As SCAN_RESULT)
    'O16 - DPF: {0000000} (shit toolbar) - http://bla.com/bla.dll
    'O16 - DPF: Plugin - http://bla.com/bla.dll
    
    On Error GoTo ErrorHandler:
    FixFileHandler result
    FixRegistryHandler result
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
    Dim UseWow, Wow6432Redir As Boolean, result As SCAN_RESULT, Data() As String, sTrimChar As String
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
                    
                    ReDim Data(0)
                    Data(0) = sData
                    
                    If sParam = "NameServer" Then
                        Data = SplitByMultiDelims(Trim$(sData), True, sTrimChar, " ", ",")
                        ArrayRemoveEmptyItems Data
                    End If
                    
                    For i = 0 To UBound(Data)
                    
                        sData = Data(i)
                    
                        sHit = "O17 - HKLM\" & IIf(j = 0, "System\CCS", CSKey) & "\" & sKeyDomain(n) & ": [" & sParam & "] = " & sData
                    
                        If sParam = "NameServer" Then
                            sProviderDNS = GetCollectionItemByKey(sData, colSafeDNS)
                            If sProviderDNS <> "" Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                        End If
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O17"
                                .HitLineW = sHit
                                'AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, CSKey & "\" & sKeyDomain(n), sParam
                                
                                AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                    HKEY_LOCAL_MACHINE, CSKey & "\" & sKeyDomain(n), sParam, _
                                    , , , CStr(Data(i)), "", sTrimChar
                                
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    
                    Next
                    
                End If
            Next
            
            'HKLM\System\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\.. subkeys
            'HKLM\System\CS*\Services\Tcpip\Parameters\Interfaces\.. subkeys
            
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
                        
                        Data = SplitByMultiDelims(Trim$(sData), True, sTrimChar, " ", ",")
                        ArrayRemoveEmptyItems Data
                        
                        For i = 0 To UBound(Data)
                            ReDim Preserve TcpIpNameServers(UBound(TcpIpNameServers) + 1)   'for using in filtering DNS DHCP later
                            TcpIpNameServers(UBound(TcpIpNameServers)) = Data(i)
                        Next
                    End If
                    
                    For i = 0 To UBound(Data)
                        
                        sHit = "O17 - HKLM\" & IIf(j = 0, "System\CCS", CSKey) & "\Services\Tcpip\..\" & aNames(n) & ": [" & sParam & "] = " & Data(i)
                        
                        If sParam = "NameServer" Then
                            sProviderDNS = GetCollectionItemByKey(CStr(Data(i)), colSafeDNS)
                            If sProviderDNS <> "" Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                        End If
                        
                        If Not IsOnIgnoreList(sHit) Then
                            With result
                                .Section = "O17"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                    HKEY_LOCAL_MACHINE, CSKey & "\Services\Tcpip\Parameters\Interfaces\" & aNames(n), sParam, _
                                    , , , CStr(Data(i)), "", sTrimChar
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
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
                sHit = IIf(bIsWin32, "O17", IIf(Wow6432Redir, "O17-32", "O17")) & " - HKLM\Software\..\Telephony: [" & sParam & "] = " & sDomain
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Section = "O17"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, sTelephonyDomain, sParam
                        .CureType = REGISTRY_BASED
                    End With
                    AddToScanResults result
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
                
                    'temporarily disabled
                    'If (Not inArray(DNS(i), TcpIpNameServers, , , vbTextCompare) Or bIgnoreAllWhitelists) Then
                        sHit = "O17 - DHCP DNS " & i + 1 & ": " & DNS(i)
                        
                        sProviderDNS = GetCollectionItemByKey(DNS(i), colSafeDNS)
                        If sProviderDNS <> "" Then sHit = sHit & " (" & "Well-known DNS: " & sProviderDNS & ")"
                        
                        If Not IsOnIgnoreList(sHit) Then AddToScanResultsSimple "O17", sHit
                    'End If
                    
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

Public Sub FixO17Item(sItem$, result As SCAN_RESULT)
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
    
    FixRegistryHandler result
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
    
    ScanPrinterPorts
    
    Dim hKey&, i&, sName$, sCLSID$, sFile$, lpcName&, sHit$, result As SCAN_RESULT
    Dim bShared As Boolean, vkt As KEY_VIRTUAL_TYPE, bSafe As Boolean, sFixKey As String
    Dim bBySubKey As Boolean, aSubKey() As String, j&, sDefCLSID As String, sDefCLSID_all As String, vDefCLSID As Variant, sDefFile As String
    Dim sHash As String
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_ALL
    HE.AddKey "Software\Classes\Protocols\Handler"
    
    Do While HE.MoveNext
        vkt = Reg.GetKeyVirtualType(HE.Hive, HE.Key)
        bShared = (vkt And KEY_VIRTUAL_SHARED)
        
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            i = 0
            Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
                sName = TrimNull(sName)
                sCLSID = UCase$(Reg.GetString(HE.Hive, HE.Key & "\" & sName, "CLSID", HE.Redirected))
                
                sFile = ""
                If sCLSID <> "" Then
                    Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, bShared)
                End If
                If sCLSID = "" Then sCLSID = "(no CLSID)"
                If sFile = "" Then sFile = "(no file)"
                
                bSafe = False
                
                If oDict.dSafeProtocols.Exists(sName) Then
                    sDefCLSID_all = oDict.dSafeProtocols(sName)
                    If bHideMicrosoft Then
                        If IsMicrosoftFile(sFile) Then
                            bSafe = True
                        Else
                            If StrComp(GetFileName(sFile, True), "MSITSS.DLL", 1) = 0 Then
                                'https://www.virustotal.com/gui/file/7941cd077397bc18e6dc46a478e196f25fd56ee0ad7ebcaede6b360c77d57de1/detection
                                'ms office 2003
                                If StrComp(GetFileSHA1(sFile, , True), "19255cb7154b30697431ec98e9c9698e39d80c7d", 1) = 0 Then bSafe = True
                            End If
                        End If
                    End If
                Else
                    sDefCLSID_all = ""
                    If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                        If bHideMicrosoft Then
                            If IsMicrosoftFile(sFile) Then bSafe = True
                        End If
                    End If
                End If
                
                If Not bSafe And Len(sFile) <> 0 Then
                    If StrComp(GetFileNameAndExt(sFile), "MSDAIPP.DLL", 1) = 0 Then
                        sHash = GetFileSHA1(sFile, , True)
                        If sHash = "3F61C6698DEA48E0CA8C05019CF470E54B4782C6" Or _
                          sHash = "A5F1FE61B58F28A65AE189AF14749DD241A9830D" Then bSafe = True
                    End If
                End If
                
                'Repeat for subkey
                
                If Not bSafe Then
                    'Protocols key can contain several subkeys with similar contents
                    bBySubKey = False
                    If sCLSID = "(no CLSID)" Then
                        If Reg.KeyHasSubKeys(HE.Hive, HE.Key & "\" & sName, HE.Redirected) Then
                            bBySubKey = True
                            For j = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key & "\" & sName, aSubKey, HE.Redirected)
                                sCLSID = UCase$(Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\" & aSubKey(j), "CLSID", HE.Redirected))
                                sFile = ""
                                If sCLSID <> "" Then
                                    Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, bShared)
                                End If
                                If sCLSID = "" Then sCLSID = "(no CLSID)"
                                If sFile = "" Then sFile = "(no file)"
                                
                                If oDict.dSafeProtocols.Exists(sName) Then
                                    sDefCLSID_all = oDict.dSafeProtocols(sName)
                                    If bHideMicrosoft Then
                                        If IsMicrosoftFile(sFile) Then bSafe = True
                                    End If
                                Else
                                    sDefCLSID_all = ""
                                    If InStr(1, sFile, "\Microsoft Office", 1) <> 0 Then
                                        If bHideMicrosoft Then
                                            If IsMicrosoftFile(sFile) Then bSafe = True
                                        End If
                                    End If
                                End If
                                
                                If Not bSafe And Len(sFile) <> 0 Then
                                    If StrComp(GetFileNameAndExt(sFile), "MSDAIPP.DLL", 1) = 0 Then
                                        sHash = GetFileSHA1(sFile, , True)
                                        If sHash = "3F61C6698DEA48E0CA8C05019CF470E54B4782C6" Or _
                                          sHash = "A5F1FE61B58F28A65AE189AF14749DD241A9830D" Then bSafe = True
                                    End If
                                End If
                                
                                If Not bSafe Then
                                    sHit = "O18 - " & HE.KeyAndHive & "\" & sName & "\" & aSubKey(j) & ": [CLSID] = " & sCLSID & " - " & sFile
                                    sFixKey = HE.Key & "\" & sName & "\" & aSubKey(j)
                                    GoSub labelFix:
                                End If
                            Next
                        End If
                    End If
                    
                    If Not bBySubKey Then
                        'HKCU often has empty protocol keys, so skip them
                        
                        'If Not (sCLSID = "(no CLSID)" And (HE.Hive = HKCU Or HE.Hive = HKU)) Then
                
                            sHit = "O18 - " & HE.KeyAndHive & "\" & sName & ": [CLSID] = " & sCLSID & " - " & sFile
                            sFixKey = HE.Key & "\" & sName
                        
                            GoSub labelFix:
                        'End If
                    End If
                End If
                
                sName = String$(MAX_KEYNAME, 0)
                lpcName = Len(sName)
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    '-------------------
    'Filters:
    
    HE.Clear
    HE.AddKey "Software\Classes\Protocols\Filter"
    
    hKey = 0
    sCLSID = vbNullString
    sFile = vbNullString
    
    Do While HE.MoveNext
    
        vkt = Reg.GetKeyVirtualType(HE.Hive, HE.Key)
        bShared = (vkt And KEY_VIRTUAL_SHARED)
        
        If RegOpenKeyExW(HE.Hive, StrPtr(HE.Key), 0, KEY_ENUMERATE_SUB_KEYS Or (bIsWOW64 And KEY_WOW64_64KEY And Not HE.Redirected), hKey) = 0 Then
            sName = String$(MAX_KEYNAME, 0&)
            lpcName = Len(sName)
            i = 0
            Do While RegEnumKeyExW(hKey, i, StrPtr(sName), lpcName, 0&, 0&, ByVal 0&, ByVal 0&) = 0
                sName = TrimNull(sName)
                sCLSID = Reg.GetString(HE.Hive, HE.Key & "\" & sName, "CLSID", HE.Redirected)
                
                If sCLSID = "" Then
                    sCLSID = "(no CLSID)"
                    sFile = "(no file)"
                Else
                    Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, bShared)
                End If
                
                bSafe = False
                If oDict.dSafeFilters.Exists(sName) Then
                    sDefCLSID_all = oDict.dSafeFilters(sName)
                    If bHideMicrosoft Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                Else
                    sDefCLSID_all = ""
                    If InStr(1, sFile, "\Microsoft Shared\", 1) <> 0 Then
                        If bHideMicrosoft Then
                            If IsMicrosoftFile(sFile) Then bSafe = True
                        End If
                    End If
                End If
                
                'Repeat for subkey
                
                If Not bSafe Then
                    'Filters key, possibly, can also contain several subkeys with similar contents
                    bBySubKey = False
                    If sCLSID = "(no CLSID)" Then
                        If Reg.KeyHasSubKeys(HE.Hive, HE.Key & "\" & sName, HE.Redirected) Then
                            bBySubKey = True
                            For j = 1 To Reg.EnumSubKeysToArray(HE.Hive, HE.Key & "\" & sName, aSubKey, HE.Redirected)
                                sCLSID = Reg.GetString(HE.Hive, HE.Key & "\" & sName & "\" & aSubKey(j), "CLSID", HE.Redirected)
                                
                                If sCLSID = "" Then
                                    sCLSID = "(no CLSID)"
                                    sFile = "(no file)"
                                Else
                                    Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, bShared)
                                End If
                                
                                If oDict.dSafeFilters.Exists(sName) Then
                                    sDefCLSID_all = oDict.dSafeFilters(sName)
                                    If bHideMicrosoft Then
                                        If IsMicrosoftFile(sFile) Then bSafe = True
                                    End If
                                Else
                                    sDefCLSID_all = ""
                                    If InStr(1, sFile, "\Microsoft Shared\", 1) <> 0 Then
                                        If bHideMicrosoft Then
                                            If IsMicrosoftFile(sFile) Then bSafe = True
                                        End If
                                    End If
                                End If
                                
                                If Not bSafe Then
                                    sHit = "O18 - " & HE.KeyAndHive & "\" & sName & ": [CLSID] = " & sCLSID & " - " & sFile
                                    sFixKey = HE.Key & "\" & sName & "\" & aSubKey(j)
                                    
                                    GoSub labelFix:
                                End If
                            Next
                        End If
                    End If
                    
                    If Not bBySubKey Then
                        'HKCU often has empty filter keys, so skip them
                        
                        'If Not (sCLSID = "(no CLSID)" And (HE.Hive = HKCU Or HE.Hive = HKU)) Then
                    
                            sHit = "O18 - " & HE.KeyAndHive & "\" & sName & ": [CLSID] = " & sCLSID & " - " & sFile
                            sFixKey = HE.Key & "\" & sName
                            
                            GoSub labelFix:
                        'End If
                    End If
                End If
                
                sName = String$(MAX_KEYNAME, 0&)
                lpcName = Len(sName)
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    AppendErrorLogCustom "CheckO18Item - End"
    Exit Sub

labelFix:
    If Not IsOnIgnoreList(sHit) Then
        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
        With result
            .Section = "O18"
            .HitLineW = sHit
            
            sDefCLSID = ""
            'find suitable CLSID by database and check is it legit
            For Each vDefCLSID In SplitSafe(sDefCLSID_all, "|")
                If Len(vDefCLSID) <> 0 Then
                    Call GetFileByCLSID(CStr(vDefCLSID), sDefFile, , HE.Redirected, bShared)
                    If IsMicrosoftFile(sDefFile) Then
                        sDefCLSID = CStr(vDefCLSID)
                        Exit For
                    End If
                End If
            Next
            If Len(sDefCLSID) = 0 Then
                AddRegToFix .Reg, REMOVE_KEY, HE.Hive, sFixKey, , , HE.Redirected
            Else
                AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, sFixKey, "CLSID", sDefCLSID, HE.Redirected
            End If

            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
            
            .CureType = REGISTRY_BASED Or FILE_BASED
        End With
        AddToScanResults result
    End If
    
    Return

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO18Item"
    If hKey <> 0 Then RegCloseKey hKey
End Sub

Public Sub ScanPrinterPorts() '// Thanks to Alex Ionescu
    On Error GoTo ErrorHandler:
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports

    Dim i&, aPorts$(), sFile$, sHit$, result As SCAN_RESULT
    
    For i = 1 To Reg.EnumValuesToArray(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports", aPorts, False)
        
        sFile = EnvironW(aPorts(i))
        
        If InStr(sFile, ":\") <> 0 Or InStr(sFile, ":/") <> 0 Then 'look as file
            
            sHit = "O18 - Printer Port: " & sFile
            
            If Not IsOnIgnoreList(sHit) Then
                
                With result
                    .Section = "O18"
                    .HitLineW = sHit
                    AddCustomToFix .Custom, CUSTOM_ACTION_SPECIFIC, aPorts(i)
                    AddRegToFix .Reg, REMOVE_VALUE, HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports", aPorts(i), , False
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ScanPrinterPorts"
End Sub

Public Function FileMissing(sFile$) As Boolean
    If Len(sFile) = 0 Then FileMissing = True: Exit Function
    If sFile = "(no file)" Then FileMissing = True: Exit Function
    If StrEndWith(sFile, "(file missing)") Then FileMissing = True: Exit Function
End Function

'Private Function O18_GetCLSIDByProtocol(sProtocol$) As String
'    Dim i&, sCLSID$
'    For i = 0 To UBound(aSafeProtocols)
'        'find CLSID for protocol name
'        If InStr(1, aSafeProtocols(i), sProtocol) > 0 Then
'            sCLSID = SplitSafe(aSafeProtocols(i), "|")(1)
'            Exit For
'        End If
'    Next i
'    O18_GetCLSIDByProtocol = sCLSID
'End Function
'
'Private Function O18_GetCLSIDByFilter(sFilter$) As String
'    Dim i&, sCLSID$
'    For i = 0 To UBound(aSafeFilters)
'        'find CLSID for protocol name
'        If InStr(1, aSafeFilters(i), sFilter) > 0 Then
'            sCLSID = SplitSafe(aSafeFilters(i), "|")(1)
'            Exit For
'        End If
'    Next i
'    O18_GetCLSIDByFilter = sCLSID
'End Function

Public Sub FixO18Item(sItem$, result As SCAN_RESULT)
    'O18 - Protocol: cn
    'O18 - Filter: text/blah - {0} - c:\file.dll
    'O18 - Printer Port: c:\file.exe
    On Error GoTo ErrorHandler:
    
    If InStr(1, result.HitLineW, "Printer Port:", 1) <> 0 Then
        
        Dim sPort As String
        sPort = result.Custom(0).ObjectName
        
        'get-printer / remove-printer are Win 8+ only?
        
        If Proc.ProcessRun(BuildPath(sWinSysDir, "WindowsPowerShell\v1.0\powershell.exe"), _
          "-ExecutionPolicy UnRestricted -c " & """" & _
          "$printer = get-printer * | where {$_.portname -eq '" & sPort & "'}; remove-printer -inputobject $printer" & """", , vbHide) Then
            Proc.WaitForTerminate , , , 15000
        End If
    End If
    
    FixRegistryHandler result
    FixFileHandler result
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
    
    Dim lUseMySS&, sUserSS$, sHit$, result As SCAN_RESULT
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    HE.Init HE_HIVE_HKCU Or HE_HIVE_HKU
    HE.AddKey "Software\Microsoft\Internet Explorer\Styles"
    
    Do While HE.MoveNext
        lUseMySS = Reg.GetDword(HE.Hive, HE.Key, "Use My Stylesheet", HE.Redirected)
        sUserSS = Reg.GetString(HE.Hive, HE.Key, "User Stylesheet", HE.Redirected)
        
        sUserSS = FormatFileMissing(sUserSS)
        
        If lUseMySS <> 0 And Len(sUserSS) <> 0 Then
            'O19 - User stylesheet (HKCU,HKLM):
            'O19-32 - User stylesheet (HKCU,HKLM):
            sHit = IIf(bIsWin32, "O19", IIf(HE.Redirected, "O19-32", "O19")) & " - " & HE.HiveNameAndSID & "\..\Internet Explorer\Styles: " & _
                "[User Stylesheet] = " & sUserSS
            If Not IsOnIgnoreList(sHit) Then
                'md5 doesn't seem useful here
                'If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sUserSS)
                With result
                    .Section = "O19"
                    .HitLineW = sHit
                    AddRegToFix .Reg, RESTORE_VALUE, HE.Hive, HE.Key, "Use My Stylesheet", 0&, HE.Redirected
                    AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, "User Stylesheet", , HE.Redirected
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
            End If
        End If
    Loop
    
    AppendErrorLogCustom "CheckO19Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO19Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO19Item(sItem$, result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    'O19 - User stylesheet: c:\file.css (file missing)
    FixRegistryHandler result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO19Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO20Item()
    'AppInit_DLLs
    'https://support.microsoft.com/ru-ru/kb/197571
    'https://msdn.microsoft.com/en-us/library/windows/desktop/dd744762(v=vs.85).aspx
    
    'According to MSDN:
    ' - modules are delimited by spaces or commas
    ' - long file names are not permitted
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO20Item - Begin"
    
    'appinit_dlls + winlogon notify
    Dim sAppInit$, sFile$, sHit$, UseWow, Wow6432Redir As Boolean, result As SCAN_RESULT
    Dim bEnabled As Boolean, bRequireCodeSigned As Boolean, aFile() As String, bUnsigned As Boolean, i As Long
    Dim sTrimChar As String, sOrigLine As String
    
    For Each UseWow In Array(False, True)
        Wow6432Redir = UseWow
        If bIsWin32 And Wow6432Redir Then Exit For
    
        sAppInit = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
        
        If OSver.MajorMinor <= 5.2 Then 'XP/2003-
            bEnabled = True
        Else
            bEnabled = (1 = Reg.GetDword(HKEY_LOCAL_MACHINE, sAppInit, "LoadAppInit_DLLs", Wow6432Redir))
            bRequireCodeSigned = (1 = Reg.GetDword(HKEY_LOCAL_MACHINE, sAppInit, "RequireSignedAppInit_DLLs", Wow6432Redir))
        End If
        
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sAppInit, "AppInit_DLLs", Wow6432Redir)
        If Len(sFile) <> 0 Then
            
            aFile = SplitByMultiDelims(sFile, True, sTrimChar, ",", " ")
            ArrayRemoveEmptyItems aFile
            
            For i = 0 To UBound(aFile)
            
                sFile = aFile(i)
                sOrigLine = sFile
                
                If (InStr(1, "*" & sSafeAppInit & "*", "*" & sFile & "*", vbTextCompare) = 0) Or bIgnoreAllWhitelists Then
                    'item is not on whitelist
                    'O20 - AppInit_DLLs
                    'O20-32 - AppInit_DLLs
                    
                    sFile = FormatFileMissing(sFile)
                    
                    If bRequireCodeSigned Then
                        bUnsigned = False
                        If FileExists(sFile) Then
                            If Not IsLegitFileEDS(sFile) Then bUnsigned = True
                        End If
                    End If
                    
                    sHit = IIf(bIsWin32, "O20", IIf(Wow6432Redir, "O20-32", "O20")) & " - HKLM\..\Windows: [AppInit_DLLs] = " & sFile & _
                      IIf(bRequireCodeSigned And bUnsigned, " (disabled because not code signed)", "") & _
                      IIf(Not bEnabled, " (disabled by registry)", "") & IIf(OSver.SecureBoot, " (disabled by SecureBoot)", "")
                    
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                    If Not IsOnIgnoreList(sHit) Then
                        With result
                            .Section = "O20"
                            .HitLineW = sHit
                            
                            'to disable loading AppInit_DLLs
                            'AddRegToFix .Reg, RESTORE_VALUE, 0, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Windows", "LoadAppInit_DLLs", 0, CLng(Wow6432Redir), REG_RESTORE_DWORD
                            
                            AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE, _
                                HKLM, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "AppInit_DLLs", , CLng(Wow6432Redir), REG_RESTORE_SZ, _
                                sOrigLine, "", sTrimChar
                            
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            Next
        End If
        
        Dim sSubkeys$(), sWinLogon$
        sWinLogon = "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify"
        sSubkeys = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, sWinLogon, Wow6432Redir), "|")
        If UBound(sSubkeys) <> -1 Then
            For i = 0 To UBound(sSubkeys)
                If (InStr(1, "*" & sSafeWinlogonNotify & "*", "*" & sSubkeys(i) & "*", vbTextCompare) = 0) Or bIgnoreAllWhitelists Then
                    sFile = Reg.GetString(HKEY_LOCAL_MACHINE, sWinLogon & "\" & sSubkeys(i), "DllName", Wow6432Redir)
                    
                    sFile = FormatFileMissing(sFile)
                    
                    'O20 - Winlogon Notify:
                    'O20-32 - Winlogon Notify:
                    sHit = IIf(bIsWin32, "O20", IIf(Wow6432Redir, "O20-32", "O20")) & " - HKLM\..\Winlogon\Notify\" & sSubkeys(i) & ": [DllName] = " & sFile
                    If Not IsOnIgnoreList(sHit) Then
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                        With result
                            .Section = "O20"
                            .HitLineW = sHit
                            AddRegToFix .Reg, REMOVE_KEY, HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify\" & sSubkeys(i), , , CLng(Wow6432Redir)
                            .CureType = REGISTRY_BASED
                        End With
                        AddToScanResults result
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

Public Sub FixO20Item(sItem$, result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    
    'O20 - AppInit_DLLs: file.dll
    'O20 - Winlogon Notify: bladibla - c:\file.dll
    '
    '* clear appinit regval (don't delete it)
    '* kill regkey (for winlogon notify)
    
    FixRegistryHandler result
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
    Dim result As SCAN_RESULT, bSafe As Boolean, bInList As Boolean
    
    sSSODL = "Software\Microsoft\Windows\CurrentVersion\ShellServiceObjectDelayLoad"
    
    'BE AWARE: SHELL32.dll - sometimes this file is patched
    '(e.g. seen after "Windown XP Update pack by Simplix" together with his certificate installed to trusted root storage)
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
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
                
                Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, HE.SharedKey)
                
                sFile = FormatFileMissing(sFile)
                
                bSafe = False
                If bHideMicrosoft And Not bIgnoreAllWhitelists Then
                    
                    bInList = inArray(sCLSID, aSafeSSODL, , , vbTextCompare)
                    
                    If bInList Then
                        If IsMicrosoftFile(sFile) Then bSafe = True
                    End If
                    If Not bSafe Then
                        If WhiteListed(sFile, "GROOVEEX.DLL", True) Then bSafe = True
                    End If
                End If
                
                If Not bSafe Then
                    Call GetTitleByCLSID(sCLSID, sName, HE.Redirected, HE.SharedKey)
                
                    If sName = "(no name)" Then sName = sValueName
                    
                    sHit = IIf(bIsWin32, "O21", IIf(HE.Redirected, "O21-32", "O21")) & " - HKLM\..\ShellServiceObjectDelayLoad: " & _
                        IIf(sName <> sValueName, sName & " ", "") & "[" & sValueName & "] " & " = " & sCLSID & " - " & sFile
                    
                    'some shit leftover by Microsoft ^)
                    If bHideMicrosoft And (sName = "WebCheck" And sCLSID = "{E6FB5E20-DE35-11CF-9C87-00AA005127ED}" And sFile = "(no file)") Then bSafe = True
                
                    If Not IsOnIgnoreList(sHit) And Not bSafe Then
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                        With result
                            .Section = "O21"
                            .HitLineW = sHit
                            If sCLSID <> "" Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                            AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sName, , HE.Redirected
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
                i = i + 1
            Loop
            RegCloseKey hKey
        End If
    Loop
    
    'ShellIconOverlayIdentifiers
    Dim aSubKey() As String
    Dim sSIOI As String
    
    sSIOI = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers"
    
    HE.Init HE_HIVE_HKLM
    HE.AddKey sSIOI
    
    Do While HE.MoveNext
        
        
        If Reg.EnumSubKeysToArray(HE.Hive, HE.Key, aSubKey, HE.Redirected) > 0 Then
        
            For i = 1 To UBound(aSubKey)
            
                sName = aSubKey(i)
                sCLSID = Reg.GetString(HE.Hive, HE.Key & "\" & aSubKey(i), vbNullString, HE.Redirected)
                
                Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, HE.SharedKey)
                
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
                
                If Not bSafe Then
                    Call GetTitleByCLSID(sCLSID, sName, HE.Redirected, HE.SharedKey)
                    
                    'If sName = "(no name)" Then sName = aSubKey(i)
                    
                    sHit = IIf(bIsWin32, "O21", IIf(HE.Redirected, "O21-32", "O21")) & " - HKLM\..\ShellIconOverlayIdentifiers\" & _
                        aSubKey(i) & ": " & sName & " - " & sCLSID & " - " & sFile
                    
                    If Not IsOnIgnoreList(sHit) And Not bSafe Then
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                        With result
                            .Section = "O21"
                            .HitLineW = sHit
                            If sCLSID <> "" Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                            AddRegToFix .Reg, REMOVE_KEY, HE.Hive, HE.Key & "\" & aSubKey(i), , , HE.Redirected
                            AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                            .CureType = REGISTRY_BASED Or FILE_BASED
                        End With
                        AddToScanResults result
                    End If
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
        
        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue, HE.Redirected)
            sCLSID = aValue(i)
            
            Call GetFileByCLSID(sCLSID, sFile, , HE.Redirected, HE.SharedKey)
            
            sFile = FormatFileMissing(sFile)
            
            bSafe = False
            If bHideMicrosoft And Not bIgnoreAllWhitelists Then
            
                bInList = inArray(sFile, aSafeSEH, , , vbTextCompare)
                If StrComp(GetFileName(sFile, True), "GROOVEEX.DLL", 1) = 0 Then bInList = True
                
                If bInList Then
                    If IsMicrosoftFile(sFile) Then bSafe = True
                End If
            End If
            
            If Not bSafe Then
                Call GetTitleByCLSID(sCLSID, sName, HE.Redirected, HE.SharedKey)
                
                If OSver.MajorMinor >= 6 Then 'XP/2003 has no policy
                    bDisabled = Not (1 = Reg.GetDword(HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "EnableShellExecuteHooks"))
                End If
                
                sHit = IIf(HE.Redirected, "O21-32", "O21") & " - HKLM\..\ShellExecuteHooks: [" & sCLSID & "] - " & sName & " - " & sFile & IIf(bDisabled, " (disabled)", "")
                If Not IsOnIgnoreList(sHit) And Not bSafe Then
                    If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                    With result
                        .Section = "O21"
                        .HitLineW = sHit
                        If sCLSID <> "" Then AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID, , , REG_REDIRECTION_BOTH
                        AddRegToFix .Reg, REMOVE_VALUE, HE.Hive, HE.Key, sCLSID, , HE.Redirected
                        AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    Loop
    
    'ColumnHandlers
    'Note: Not available in Vista+
    'HKEY_CLASSES_ROOT\Folder\ShellEx\ColumnHandlers
    
    'see also: https://www.nirsoft.net/utils/shexview.html
    
'    HE.Init HE_HIVE_ALL
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\ColumnHandlers"
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\ContextMenuHandlers"
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\DragDropHandlers"
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\PropertySheetHandlers"
'    HE.AddKey "SOFTWARE\Classes\Folder\ShellEx\CopyHookHandlers"
    
    'Нужно подумать, как это представить в логе. Обработчики могут быть зарегистрированы как для папок, так и для отдельных расширений имени.
    'Сканировать всё подряд? Нужно будет совмещать одинаковые CLSID.
    
    'HKEY_CLASSES_ROOT\*\shellex
    'HKEY_CLASSES_ROOT\Folder\shellex - virtual combination of actual file folders, shell folders, drives, other special folders
    'HKEY_CLASSES_ROOT\Directory\shellex
    'and so ...
    
    'Explorer Shell extensions
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved
    'see:
    'https://msdn.microsoft.com/en-us/library/ms812054.aspx
    'https://forum.sysinternals.com/shell-extensions-approved_topic11891.html
    'So, this key can be needed only for heuristic cleaning of extension
    
    
    AppendErrorLogCustom "CheckO21Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO21Item"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO21Item(sItem$, result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    
    'O21 - SSODL: webcheck - {000....000} - c:\file.dll (file missing)
    'actions to take:
    '* kill file
    '* kill regkey - ShellIconOverlayIdentifiers
    '* kill regparam - SSODL
    '* kill clsid regkey
    
    ShutdownExplorer
    FixRegistryHandler result
    FixFileHandler result
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
    
    'Win XP / Server 2003
    
    Dim sSTS$, hKey&, i&, sCLSID$, lCLSIDLen&, lDataLen&
    Dim sFile$, sName$, sHit$, isSafe As Boolean
    Dim Wow6432Redir As Boolean, result As SCAN_RESULT
    
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
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                With result
                    .Section = "O22"
                    .HitLineW = sHit
                    AddRegToFix .Reg, REMOVE_VALUE, HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\SharedTaskScheduler", sCLSID
                    AddRegToFix .Reg, REMOVE_KEY, HKEY_CLASSES_ROOT, "CLSID\" & sCLSID
                    AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sFile
                    .CureType = REGISTRY_BASED
                End With
                AddToScanResults result
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

Public Sub FixO22Item(sItem$, result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    'O22 - ScheduledTask: blah - {000...000} - file.dll
    FixIt result
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
    Dim sServices$(), i&, j&, k&, sName$, sDisplayName$, tmp$, result As SCAN_RESULT
    Dim lStart&, lType&, sFile$, sHit$, sBuf$, IsCompositeCmd As Boolean
    Dim bHideDisabled As Boolean, sServiceDll As String, sServiceDll_2 As String, bDllMissing As Boolean
    Dim ServState As SERVICE_STATE
    Dim argc As Long
    Dim argv() As String
    Dim isSafeMSCmdLine As Boolean
    'Dim SignResult As SignResult_TYPE
    Dim bMicrosoft As Boolean
    Dim FoundFile As String
    Dim IsMSCert As Boolean
    Dim sImagePath As String
    Dim sArgument As String
    Dim pos As Long
    Dim bSuspicious As Boolean
    Dim sGroup As String
    
    Dim dLegitService As clsTrickHashTable
    Set dLegitService = New clsTrickHashTable
    dLegitService.CompareMode = vbTextCompare
    
    Dim dLegitGroups As clsTrickHashTable
    Set dLegitGroups = New clsTrickHashTable
    dLegitGroups.CompareMode = vbTextCompare
    
    If Not bIsWinNT Then Exit Sub
    
    If Not bIgnoreAllWhitelists Then
        bHideDisabled = True
    End If
    
    sServices = Split(Reg.EnumSubKeys(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services"), "|")
    
    If UBound(sServices) = -1 Then Exit Sub
    
    For i = 0 To UBound(sServices)
        
        sName = sServices(i)
        Dbg sName
        
        'If InStr(sName, "EventSystem2") Then Stop
        
        lType = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Type")
        
        If lType < 16 Then 'Driver
            If Not bAdditional Then 'if 'O23 - Driver' check is skipped
                If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 4&
                sGroup = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Group")
                If Len(sGroup) <> 0 Then
                    If Not dLegitGroups.Exists(sGroup) Then dLegitGroups.Add sGroup, 4&
                End If
            End If
            GoTo Continue
        End If
        
        lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start")
        
        If (lStart = 4 And bHideDisabled) Then
            If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 4&
            GoTo Continue
        End If
        
        UpdateProgressBar "O23", sName
        
        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ImagePath")
        sServiceDll = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName & "\Parameters", "ServiceDll")
        
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
        sImagePath = sFile
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
        
        sArgument = ""
        bSuspicious = False
        
        If lType >= 16 Then
          If Not (lStart = 4 And bHideDisabled) Then
            
            IsCompositeCmd = False
            isSafeMSCmdLine = False
            
            'признак командной строки - вся строка как файл не существует на диске
            If Not FileExists(sImagePath) And sImagePath <> "" Then
            
                ' Дальше идут процедуры парсинга командной строки и проверки сертиката для каждого файла из этой цепочки
                ' Если любой файл из цепочки не проходит проверку, строка считается небезопасной
            
                ParseCommandLine sImagePath, argc, argv
                
                If argc > 1 Then
                    pos = InStr(sImagePath, argv(1))
                    If pos <> 0 Then
                        sArgument = Mid$(sImagePath, pos + Len(argv(1)))
                        If Left$(sArgument, 1) = """" Then sArgument = Mid$(sArgument, 2)
                        sArgument = LTrim$(sArgument)
                    End If
                End If
                
                If Len(sArgument) <> 0 Then
                    If Len(sArgument) > 50 Then
                        bSuspicious = True
                    ElseIf InStr(1, sArgument, "http", 1) <> 0 Then
                        bSuspicious = True
                    ElseIf InStr(1, EnvironW(sArgument), "http", 1) <> 0 Then
                        bSuspicious = True
                    End If
                End If
                
                '// TODO: добавить к FindOnPath папку, в которой находится основной запускаемый службой файл
                
                'если файл в составе коммандной строки, например: C:\WINDOWS\system32\svchost -k rpcss.exe
                
                If argc > 2 Then        ' 1 -> app exe self, 2 -> actual cmd, 3 -> arg
                
                  If Not FileExists(argv(1)) Then   ' если запускающий файл не существует -> ищем его
                    FoundFile = FindOnPath(argv(1))
                    argv(1) = FoundFile
                  Else
                    FoundFile = argv(1)
                  End If
                
                  ' если запускающий файл существует (иначе, нет смысла проверять остальные аргументы)
                  If 0 <> Len(FoundFile) And Not StrBeginWith(sImagePath, sWinSysDir & "\svchost.exe -k") Then
                  
                    'флаг о том, что служба запускает составную командную строку, в которой как минимум первый (запускающий файл) существует
                    IsCompositeCmd = True
                
                    isSafeMSCmdLine = True
                 
                    For j = 1 To UBound(argv) ' argv[1] -> запускающий файл в цепочке
                    
                        ' проверяем хеш корневого сертификата каждого из элементов командной строки, если он был найден по известным путям Path
                        
                        FoundFile = FindOnPath(argv(j))
                        
                        If 0 <> Len(FoundFile) Then
                        
                            If IsWinServiceFileName(FoundFile) Then
                                'SignVerify FoundFile, SV_LightCheck, SignResult
                                'IsMSCert = SignResult.isMicrosoftSign And SignResult.isLegit
                                
                                IsMSCert = IsMicrosoftFile(FoundFile)
                            Else
                                IsMSCert = False
                            End If
                            
                            If Not IsMSCert Then isSafeMSCmdLine = False: Exit For
                        End If
                    Next
                  End If
                End If
            
            End If
            
            If 0 = Len(sFile) Then
                sFile = "(no file)"
            Else
                If (Not FileExists(sFile)) And (Not IsCompositeCmd) Then
                    sFile = sFile & " (file missing)"
                Else
'                    If IsCompositeCmd Then
'                        FoundFile = argv(1)
'                    Else
'                        FoundFile = sFile
'                    End If
'                    Stady = 33: Dbg CStr(Stady)
                    
                    'sCompany = GetFilePropCompany(FoundFile)
                    'If Len(sCompany) = 0 Then sCompany = "Unknown owner"
                    
                End If
            End If
            
            'WipeSignResult SignResult
            
            bMicrosoft = False
            
            If IsCompositeCmd Then
                If Not isSafeMSCmdLine Then bSuspicious = True
            Else
                If sFile <> "(no file)" Then    'иначе, такая проверка уже выполнена ранее
                    If IsWinServiceFileName(sFile, sArgument) Then
                        'SignVerify sFile, SV_LightCheck, SignResult
                        bMicrosoft = IsMicrosoftFile(sFile)
                    Else
                        'WipeSignResult SignResult
                    End If
                End If
            End If
            
            'override by checkind EDS of service dll if original file is Microsoft (usually, svchost)
            If bDllMissing Then
                bMicrosoft = False
            Else
                If Len(sServiceDll) <> 0 Then
                    If IsWinServiceFileName(sServiceDll) Then
                        'SignVerify sServiceDll, SV_LightCheck, SignResult
                        bMicrosoft = IsMicrosoftFile(sServiceDll)
                    Else
                        'WipeSignResult SignResult
                        bMicrosoft = False
                    End If
                End If
            End If
            
            'With SignResult
                'добавляем в список легитимных служб для дальнейшего использования при проверке зависимостей
                'If Not (bSuspicious Or bDllMissing Or Not (.isMicrosoftSign And .isLegit)) Then
                If Not (bSuspicious Or bDllMissing Or Not (bMicrosoft)) Then
                    If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 0&
                End If
                
                ' если корневой сертификат цепочки доверия принадлежит Майкрософт + с учётом, что файл проходит по базе, то исключаем службу из лога
                
                'If bSuspicious Or bDllMissing Or Not (.isMicrosoftSign And .isLegit And bHideMicrosoft) Then
                If bSuspicious Or bDllMissing Or Not (bMicrosoft And bHideMicrosoft) Then
                    
                    sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DisplayName")

                    If Len(sDisplayName) = 0 Then
                        sDisplayName = sName
                    Else
                        If Left$(sDisplayName, 1) = "@" Then                    'extract string resource from file

                            sBuf = GetStringFromBinary(, , sDisplayName)

                            If 0 <> Len(sBuf) Then sDisplayName = sBuf
                        End If
                    End If
                    
'                    pos = InStr(1, sDisplayName, "; {PlaceHolder", 1)
'                    If pos <> 0 Then
'                        sDisplayName = UnQuote(Trim$(Left$(sDisplayName, pos - 1)))
'                    End If
                    
                    sHit = "O23 - Service " & IIf(ServState <> SERVICE_STOPPED, "R", "S") & lStart & _
                        ": " & IIf(sDisplayName = sName, sName, sDisplayName & " - (" & sName & ")") & " - " & ConcatFileArg(sFile, sArgument)
                    
                    If Len(sServiceDll) = 0 Then
                        If g_bCheckSum Then
                            If sFile <> "(no file)" Then sHit = sHit & GetFileCheckSum(sFile)
                        End If
                    Else
                        sHit = sHit & "; ""ServiceDll"" = " & sServiceDll
                        If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sServiceDll)
                    End If
                    
' I temporarily remove EDS name in log
'                    If .isLegit And 0 <> Len(.SubjectName) And Not bDllMissing Then
'                        sHit = sHit & " (" & .SubjectName & ")"
'                    Else
'                        sHit = sHit & " (not signed)"
'                    End If
                    
                    If Not IsOnIgnoreList(sHit) Then
                        
                        With result
                            .Section = "O23"
                            .HitLineW = sHit
                            .Name = sName 'used in "Disable" stuff
                            .State = IIf(lStart <> 4, ITEM_STATE_ENABLED, ITEM_STATE_DISABLED)
                            
                            AddServiceToFix .Service, DELETE_SERVICE Or USE_FEATURE_DISABLE, sName, , , , ServState
                        
                            If Len(sServiceDll) = 0 Then
                                'AddJumpFile .Jump, JUMP_FILE, sFile
                                AddFileToFix .File, BACKUP_FILE, sFile
                            Else
                                'AddJumpFile .Jump, JUMP_FILE, sServiceDll
                                AddFileToFix .File, BACKUP_FILE, sServiceDll
                            End If
                            
                            'AddJumpRegistry .Jump, JUMP_KEY, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName
                            AddRegToFix .Reg, BACKUP_KEY, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName
                            .Reboot = True
                            .CureType = SERVICE_BASED Or FILE_BASED Or REGISTRY_BASED
                        End With
                        AddToScanResults result
                    End If
                End If
            'End With
          End If
        End If
Continue:
    Next i
    
    'checking drivers
    
    UpdateProgressBar "O23-D"
    
    If bAdditional Then
        CheckO23Item_Drivers sServices, dLegitService
    End If
    
    'Check dependency *(should go after 'O23 - Drivers' scan !!!)
    
    'Temporarily added to "Additional scan", until I figure out all cases with damaged EDS subsystem
    If bAdditional Then
        CheckO23Item_Dependency sServices, dLegitService, dLegitGroups
    End If
    
    Set dLegitService = Nothing
    
    AppendErrorLogCustom "CheckO23Item - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_CheckO23Item", "Service=", sDisplayName
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO23Item_Dependency(sServices() As String, dLegitService As clsTrickHashTable, dLegitGroups As clsTrickHashTable)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO23Item_Dependency - Begin"
    
    Dim i&, k&
    Dim aDepend()       As String
    Dim vName           As Variant
    Dim sName           As String
    Dim sGroup          As String
    Dim sHit            As String
    Dim bMissing        As Boolean
    Dim bSafe           As Boolean
    Dim result          As SCAN_RESULT
    Dim aSubKey()       As String
    
    '"DependOnService" parameter
    
    UpdateProgressBar "O23", "Dependency"
    
    'Appending list of legit services with services that have no "ImagePath" (XP/2003- only)
    If OSver.MajorMinor <= 5.2 Then
        For i = 1 To Reg.EnumSubKeysToArray(HKLM, "System\CurrentControlSet\Services", aSubKey)
            If Not Reg.ValueExists(HKLM, "System\CurrentControlSet\Services\" & aSubKey(i), "ImagePath") Then
                If Not dLegitService.Exists(aSubKey(i)) Then
                    dLegitService.Add aSubKey(i), 0&
                End If
            End If
        Next
    End If
    
    For Each vName In dLegitService.Keys
        
        If dLegitService(vName) <> 4 Then 'check only real legit services (4 - is unknown state)
        
            sName = vName
            
            aDepend = Reg.GetMultiSZ(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DependOnService")
    
            If AryItems(aDepend) Then
                For k = 0 To UBound(aDepend)
                    
                    If Not dLegitService.Exists(aDepend(k)) Then
                        
                        bMissing = Not Reg.KeyExists(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & aDepend(k))
                        
                        bSafe = False
                        
                        'Win10 bug fix
                        'See: https://answers.microsoft.com/en-us/windows/forum/windows_10-other_settings/dependonservice-refers-to-a-non-existent-service/f63265d1-70ee-4561-b473-e54085cdeaf2
                        If bMissing And OSver.MajorMinor = 10 And Not bIgnoreAllWhitelists Then
                            If StrComp(aDepend(k), "UcmCx", 1) = 0 Then bSafe = True
                            If StrComp(aDepend(k), "GPIOClx", 1) = 0 Then bSafe = True
                        End If
                        
                        If Not bSafe Then
                        
                            sHit = "O23 - Dependency: Microsoft Service '" & sName & "' depends on unknown service: '" & aDepend(k) & "'" & _
                                IIf(bMissing, " (service missing)", "")
                            
                            If Not IsOnIgnoreList(sHit) Then
                                
                                With result
                                    .Section = "O23"
                                    .HitLineW = sHit
                                    
                                    AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                        HKLM, "System\CurrentControlSet\Services\" & sName, "DependOnService", , , REG_RESTORE_MULTI_SZ, _
                                        aDepend(k), "", vbNullChar
                                    
                                    .Reboot = True
                                    .CureType = REGISTRY_BASED
                                End With
                                AddToScanResults result
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    'Check dependency on groups
    
    '"DependOnGroup" parameter
    
    'Note: Serice group can be created by specifying "Group" registry parameter for some service.
    'There is no separate list.
    'Service groups loading order is stored in:
    ' - HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\GroupOrderList
    ' - HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\ServiceGroupOrder [List]
    ' - "Tag" reg. parameter - is an order of service loading in particular service group
    
    'Groups can be:
    '1. Legit (all services of this group belong to Microsoft)
    '2. Semi-legit (group contains both Microsoft and non-Microsoft services)
    '3. Non-legit (group contains non-Microsoft services only)
    
    'Firstly, we'll add legit and semi-legit groups to dLegitGroups dictionary.
    'And compare "Group" reg. parameter of each service with this dictionary.
    'Next, we'll list all semi-legit group to HJT log because wtf, that is wrong.
    
    'Gather groups
    
    For i = 0 To UBound(sServices)
        
        sName = sServices(i)
        sGroup = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Group")
        
        'idx description
        '1 - legit
        '2 - semi-legit
        '3 - non-legit
        '4 - state is unknown, because check is not performed. It can be due:
        ' > "Additional scan" is not performed (so, no Driver services check)
        ' > Service state is "disabled" (so such services is also not checked), unless "Ignore ALL Whitelist" is marked
        
        If Len(sGroup) <> 0 Then
            If dLegitGroups.Exists(sGroup) Then
            
                k = dLegitGroups(sGroup)
                
                If dLegitService.Exists(sName) Then
                    'Current is Legit: Make 3 => 2
                    If k = 3 Then dLegitGroups(sGroup) = 2
                Else
                    'Current is Non-legit: Make 1 => 2
                    If k = 1 Then dLegitGroups(sGroup) = 2
                End If
            Else
                If dLegitService.Exists(sName) Then
                    dLegitGroups.Add sName, 1 'make legit
                Else
                    dLegitGroups.Add sName, 3 'make non-legit
                End If
            End If
        End If
    Next
    'We should have only legit and semi-legit groups
    'So remove non-legit:
    For Each vName In dLegitGroups.Keys
        If dLegitGroups(vName) = 3 Then dLegitGroups.Remove vName
    Next
    
    'Ok, now we'll check all entries in "DependOnGroup" parameter of legit. services against "dLegitGroups" dictionary
    
    For Each vName In dLegitService.Keys
    
        If dLegitService(vName) <> 4 Then 'check only real legit services (4 - is unknown state)
        
            sName = vName
            
            aDepend = Reg.GetMultiSZ(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DependOnGroup")
            
            If AryItems(aDepend) Then
                For k = 0 To UBound(aDepend)
                    
                    If Not dLegitGroups.Exists(aDepend(k)) Then
                        
                        sHit = "O23 - Dependency: Microsoft Service '" & sName & "' depends on mixed group: '" & aDepend(k) & "'"
                        
                        If Not IsOnIgnoreList(sHit) Then
                            
                            With result
                                .Section = "O23"
                                .HitLineW = sHit
                                
                                AddRegToFix .Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                    HKLM, "System\CurrentControlSet\Services\" & sName, "DependOnGroup", , , REG_RESTORE_MULTI_SZ, _
                                    aDepend(k), "", vbNullChar
                                
                                .Reboot = True
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    'And the last: we'll list all semi-legit groups ("Additional scan" only) - can contain legit entries!
    
    If bAdditional Then
      For Each vName In dLegitGroups.Keys
        
        If dLegitGroups(vName) = 2 Then
            
            sGroup = vName
            
            'get services that belong to it
            For i = 0 To UBound(sServices)
                If StrComp(sGroup, Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sServices(i), "Group"), 1) = 0 Then
                    If Not dLegitService.Exists(sServices(i)) Then
                    
                        sHit = "O23 - Dependency: Microsoft Service Group '" & sGroup & "' contains unknown service:  '" & sServices(i) & "'"
                        
                        If Not IsOnIgnoreList(sHit) Then
                            
                            With result
                                .Section = "O23"
                                .HitLineW = sHit
                                AddRegToFix .Reg, REMOVE_VALUE, HKLM, "System\CurrentControlSet\Services\" & sServices(i), "Group"
                                .Reboot = True
                                .CureType = REGISTRY_BASED
                            End With
                            AddToScanResults result
                        End If
                    End If
                End If
            Next
        End If
      Next
    End If
    
    Set dLegitGroups = Nothing
    
    AppendErrorLogCustom "CheckO23Item_Dependency - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO23Item_Dependency"
    If inIDE Then Stop: Resume Next
End Sub

Public Sub CheckO23Item_Drivers(sServices() As String, dLegitService As clsTrickHashTable)
    'https://www.bleepingcomputer.com/tutorials/how-malware-hides-as-a-service/
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "CheckO23Item_Drivers - Begin"
    
    'Device drivers
    
    'Похоже, нужно дорабатывать процесс проверки ЭЦП
    'У драйверов Microsoft в основном идёт подпись только по сертификатам, внутренней нет.
    'У остальных - стоит кросс-подпись Microsoft, следовательно её нужно отфильтровать, и смотреть есть ли сторонняя подпись.
    'Если есть, то выводить в лог + её получателя.
    '
    '+ нужно определяться, каким методом производить удаление драйвера.
    'Судя по анализу разницы между логами, полученными через NtQuerySystemInformation и чтение реестра,
    'по всей видимости некоторые из драйверов подгрузили другие драйвера, и в этом случае непонятно как их удалять.
    'Т.е. для этих записей удаление через ветку служб отпадает.
    
    'Да и вообще, есть ли смысл получать список драйверов с помощью программы, у которой нет механизма антируткита ???
    
    ' Uninstall Devices
    '
    'https://docs.microsoft.com/en-us/windows-hardware/drivers/install/using-setupapi-to-uninstall-devices-and-driver-packages
    'https://stackoverflow.com/questions/12756712/windows-device-uninstall-using-c
    '
    ' Uninstall Drivers
    '
    'http://www.cyberforum.ru/drivers-programming/thread1300444.html#post6857698
    
    'List of mapped driver filenames:
    
    Dim dMapped As clsTrickHashTable
    Set dMapped = New clsTrickHashTable
    dMapped.CompareMode = vbTextCompare
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_DUMPFVE.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_DISKDUMP.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_ATAPI.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_DUMPATA.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_IASTORA.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_MSAHCI.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_STORAHCI.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_AMDSATA.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "DRIVERS\DUMP_AMD_SATA.SYS"), 0&
    dMapped.Add BuildPath(sWinSysDir, "Drivers\dump_pvscsi.sys"), 0&
    dMapped.Add BuildPath(sWinSysDir, "Drivers\dump_vmscsi.sys"), 0&
    dMapped.Add BuildPath(sWinSysDir, "Drivers\dump_megasas.sys"), 0&
    dMapped.Add BuildPath(sWinSysDir, "Drivers\DUMP_MEGASAS2.sys"), 0&
    dMapped.Add BuildPath(sWinSysDir, "Drivers\dump_LSI_SCSI.sys"), 0& 'Win10
    dMapped.Add BuildPath(sWinSysDir, "Drivers\dump_LSI_SAS.sys"), 0&  'Win10
    dMapped.Add BuildPath(sWinSysDir, "Drivers\dump_WMILIB.SYS"), 0&  'WinXP
    dMapped.Add BuildPath(sWinSysDir, "Drivers\DUMP_FTOIIS.SYS"), 0&
    
    'Enum Drivers via NtQuerySystemInformation:
    
    Const DRIVER_INFORMATION            As Long = 11
    Const SYSTEM_MODULE_SIZE            As Long = 284
    Const STATUS_INFO_LENGTH_MISMATCH   As Long = &HC0000004
    
    'temporarily disabled until I figure out how correctly set all filters for Microsoft entries
    
    Dim ret             As Long
    Dim buf()           As Byte
    Dim mdl             As SYSTEM_MODULE_INFORMATION
    Dim dDriver         As clsTrickHashTable
    Dim sFile           As String
    Dim i               As Long
    Dim sName           As String
    Dim lType           As Long
    Dim lStart          As Long
    Dim sDisplayName    As String
    Dim bHideDisabled   As Boolean
    Dim sBuf            As String
    Dim ServState       As SERVICE_STATE
    Dim sHit            As String
    Dim result          As SCAN_RESULT
    Dim bSafe           As Boolean
    
    If Not bIgnoreAllWhitelists Then
        bHideDisabled = True
    End If
    
    Set dDriver = New clsTrickHashTable
    dDriver.CompareMode = TextCompare
    
    If NtQuerySystemInformation(DRIVER_INFORMATION, ByVal 0&, 0, ret) = STATUS_INFO_LENGTH_MISMATCH Then
        ReDim buf(ret - 1)
        If NtQuerySystemInformation(DRIVER_INFORMATION, buf(0), ret, ret) = STATUS_SUCCESS Then
            mdl.ModulesCount = buf(0) Or (buf(1) * &H100&) Or (buf(2) * &H10000) Or (buf(3) * &H1000000)
            If mdl.ModulesCount Then
                ReDim mdl.Modules(mdl.ModulesCount - 1)
                For ret = 0 To mdl.ModulesCount - 1
                    memcpy mdl.Modules(ret), buf(ret * SYSTEM_MODULE_SIZE + 4), SYSTEM_MODULE_SIZE
                    sFile = TrimNull(mdl.Modules(ret).Name)

                    sFile = CleanServiceFileName(sFile, "")
                    
                    sFile = FindOnPath(sFile, True, sWinSysDir & "\Drivers")
                    
                    UpdateProgressBar "O23-D", sFile

                    If Not IsMicrosoftDriverFile(sFile) Or Not bHideMicrosoft Then
                        dDriver.Add sFile, 0&
                    End If

'                    If Not IsMicrosoftFile(sFile) Or Not bHideMicrosoft Then
'
'                        If InStr(1, sFile, "amtqkbgr.SYS", 1) <> 0 Then Stop
'
'                        sFile = FormatFileMissing(sFile)
'
'                        sHit = "O23 - Driver: " & sFile
'
'                        If Not IsOnIgnoreList(sHit) Then
'
'                            With Result
'                                .Section = "O23"
'                                .HitLineW = sHit
'                                AddServiceToFix .Service, DELETE_SERVICE, sName
'                                .CureType = SERVICE_BASED
'                            End With
'                            AddToScanResults Result
'                        End If
'                    End If
                Next
            End If
        End If
    End If
    
    'Enum Drivers via Registry
    
    For i = 0 To UBound(sServices)

        sName = sServices(i)

        lType = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Type")

        If lType >= 16 Then 'not a Driver
            GoTo Continue2
        End If

        lStart = Reg.GetDword(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "Start")

        If (lStart = 4 And bHideDisabled) Then
            If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 4&
            GoTo Continue2
        End If

        sDisplayName = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "DisplayName")
        
        If Len(sDisplayName) = 0 Then
            sDisplayName = sName
        Else
            If Left$(sDisplayName, 1) = "@" Then                    'extract string resource from file

                sBuf = GetStringFromBinary(, , sDisplayName)

                If 0 <> Len(sBuf) Then sDisplayName = sBuf
            End If
        End If
        
        UpdateProgressBar "O23-D", sDisplayName

        sFile = Reg.GetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName, "ImagePath")
        If Len(sFile) = 0 Then GoTo Continue2

        sFile = CleanServiceFileName(sFile, sName)

        ServState = GetServiceRunState(sName)

        If Not bAutoLogSilent Then DoEvents

        bSafe = IsMicrosoftDriverFile(sFile)

        If bSafe Then
            If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 0&
        End If
        
        If Not bSafe Or Not bHideMicrosoft Then
            
            If dDriver.Exists(sFile) Then dDriver.Remove sFile
            
            sFile = FormatFileMissing(sFile)

            sHit = "O23 - Driver " & IIf(ServState <> SERVICE_STOPPED, "R", "S") & lStart & _
                ": " & IIf(sDisplayName = sName, sName, sDisplayName & " - (" & sName & ")") & " - " & sFile
            
            If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)

            If Not IsOnIgnoreList(sHit) Then

                With result
                    .Section = "O23"
                    .HitLineW = sHit
                    AddServiceToFix .Service, DELETE_SERVICE Or USE_FEATURE_DISABLE, sName, , , , ServState
                    AddFileToFix .File, BACKUP_FILE, sFile
                    'AddJumpRegistry .Jump, JUMP_KEY, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName
                    'AddJumpFile .Jump, JUMP_FILE, sFile
                    AddRegToFix .Reg, BACKUP_KEY, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & sName
                    .Reboot = True
                    .CureType = SERVICE_BASED Or FILE_BASED Or REGISTRY_BASED
                End With
                AddToScanResults result
            Else
                If Not dLegitService.Exists(sName) Then dLegitService.Add sName, 0&
            End If
        End If

Continue2:
    Next
    
    'if currently there are more running drivers (e.g. loaded dynamically)
    If dDriver.Count > 0 Then
        For i = 0 To dDriver.Count - 1
        
            sFile = dDriver.Keys(i)
            
            bSafe = False
            'skip Microsoft drivers mapped to non-existent filename
            If Not FileExists(sFile) Then
                If dMapped.Exists(sFile) Then
                    bSafe = True
                End If
            End If
            
            If Not bSafe Or bIgnoreAllWhitelists Then
                
                sDisplayName = GetFileProperty(sFile, "FileDescription")
                If Len(sDisplayName) = 0 Then
                    sDisplayName = GetFileProperty(sFile, "ProductName")
                End If
                
                sFile = FormatFileMissing(sFile)
                
                sHit = "O23 - Driver R: " & IIf(sDisplayName = "", "(no name)", sDisplayName) & " - " & sFile
                
                If g_bCheckSum Then sHit = sHit & GetFileCheckSum(sFile)
                
                If Not IsOnIgnoreList(sHit) Then
                    
                    With result
                        .Section = "O23"
                        .HitLineW = sHit
                        AddFileToFix .File, REMOVE_FILE, sFile
                        .CureType = FILE_BASED
                        .Reboot = True
                    End With
                    AddToScanResults result
                End If
            End If
        Next
    End If
    
    AppendErrorLogCustom "CheckO23Item_Drivers - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckO23Item_Drivers", "Service=", sDisplayName
    If inIDE Then Stop: Resume Next
End Sub
    

Public Function IsWinServiceFileName(sFilePath As String, Optional sArgument As String) As Boolean
    
    On Error GoTo ErrorHandler:
    
    Static isInit As Boolean
    Dim sCompany As String
    Dim sFilename As String
    Dim sArgBase As String
    
    If Not isInit Then
        Dim vKey, prefix$
        isInit = True
        Set oDictSRV = New clsTrickHashTable
        
        'Note: this list is used to improve scan speed
        
        With oDictSRV
            .CompareMode = TextCompare
            .Add "<PF32>\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe", 0&
            .Add "<PF32>\Common Files\Microsoft Shared\OFFICE12\ODSERV.EXE", 0&
            .Add "<PF32>\Common Files\Microsoft Shared\Phone Tools\CoreCon\11.0\bin\IpOverUsbSvc.exe", 0&
            .Add "<PF32>\Common Files\Microsoft Shared\Source Engine\OSE.exe", 0&
            .Add "<PF32>\Common Files\Microsoft Shared\VS7DEBUG\MDM.exe", 0&
            .Add "<PF32>\Microsoft Office\Office12\GrooveAuditService.exe", 0&
            .Add "<PF32>\Microsoft Office\Office14\GROOVE.EXE", 0&
            .Add "<PF32>\Microsoft Application Virtualization Client\sftvsa.exe", 0&
            .Add "<PF32>\Microsoft Visual Studio\Shared\Common\DiagnosticsHub.Collection.Service\StandardCollector.Service.exe", 0&
            .Add "<PF32>\Windows Kits\8.1\App Certification Kit\fussvc.exe", 0&
            .Add "<PF32>\Skype\Updater\Updater.exe", 0&
            .Add "<PF32>\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe", 0&
            .Add "<PF64>\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe", 0&
            .Add "<PF64>\Common Files\Microsoft Shared\OFFICE12\ODSERV.EXE", 0&
            .Add "<PF64>\Common Files\Microsoft Shared\OfficeSoftwareProtectionPlatform\OSPPSVC.exe", 0&
            .Add "<PF64>\Common Files\Microsoft Shared\Windows Live\WLIDSVC.EXE", 0&
            .Add "<PF64>\Microsoft Office\Office12\GrooveAuditService.exe", 0&
            .Add "<PF64>\Microsoft Office\Office14\GROOVE.EXE", 0&
            .Add "<PF64>\Microsoft SQL Server\90\Shared\sqlwriter.exe", 0&
            .Add "<PF64>\rempl\sedsvc.exe", 0&
            .Add "<PF64>\Windows Live\Mesh\wlcrasvc.exe", 0&
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
            .Add "<SysRoot>\system32\spool\drivers\W32X86\3\PrintConfig.dll", 0&
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
            .Add "<SysRoot>\System32\msiexec.exe", "/V"
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
            .Add "<SysRoot>\System32\osrss.dll", 0&
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
            .Add "<SysRoot>\System32\SgrmBroker.exe", 0&
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
            .Add "<SysRoot>\System32\wuaueng2.dll", 0&
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
            .Add "<SysRoot>\Microsoft.NET\Framework64\v4.0.30319\SMSvcHost.exe", 0&
            .Add "<SysRoot>\system32\lserver.exe", 0&
            .Add "<SysRoot>\system32\mprdim.dll", 0&
            .Add "<SysRoot>\system32\wdfmgr.exe", 0&
            .Add "<SysRoot>\system32\sacsvr.dll", 0&
            .Add "<SysRoot>\system32\RSoPProv.exe", 0&
            .Add "<SysRoot>\system32\Dfssvc.exe", 0&
            .Add "<SysRoot>\system32\ntfrs.exe", 0&
            .Add "<SysRoot>\System32\dusmsvc.dll", 0&
            .Add "<SysRoot>\system32\SEMgrSvc.dll", 0&
            .Add "<SysRoot>\System32\SshBroker.dll", 0&
            .Add "<SysRoot>\System32\SshProxy.dll", 0&
            .Add "<SysRoot>\System32\TokenBroker.dll", 0&
            .Add "<SysRoot>\System32\debugregsvc.dll", 0&
            .Add "<SysRoot>\System32\assignedaccessmanagersvc.dll", 0&
            .Add "<SysRoot>\system32\CapabilityAccessManager.dll", 0&
            .Add "<SysRoot>\System32\DeveloperToolsSvc.exe", 0&
            .Add "<SysRoot>\System32\DevicesFlowBroker.dll", 0&
            .Add "<SysRoot>\system32\DiagSvc.dll", 0&
            .Add "<SysRoot>\System32\GraphicsPerfSvc.dll", 0&
            .Add "<SysRoot>\System32\IpxlatCfg.dll", 0&
            .Add "<SysRoot>\System32\lpasvc.dll", 0&
            .Add "<SysRoot>\System32\NaturalAuth.dll", 0&
            .Add "<SysRoot>\System32\PrintWorkflowService.dll", 0&
            .Add "<SysRoot>\System32\SharedRealitySvc.dll", 0&
            .Add "<SysRoot>\System32\Windows.WARP.JITService.dll", 0&
            .Add "<SysRoot>\System32\wfdsconmgrsvc.dll", 0&
            .Add "<SysRoot>\system32\spectrum.exe", 0&
            .Add "<SysRoot>\system32\PushToInstall.dll", 0&
            .Add "<SysRoot>\system32\InstallService.dll", 0&
            .Add "<SysRoot>\System32\XboxGipSvc.dll", 0&
            .Add "<SysRoot>\system32\xbgmsvc.exe", 0&
            .Add "<SysRoot>\System32\dns.exe", 0&
            .Add "<SysRoot>\System32\wins.exe", 0&
            .Add "<SysRoot>\System32\WBEM\WinMgmt.exe", 0&
            .Add "<SysRoot>\system32\RsSub.exe", 0&
            .Add "<SysRoot>\system32\RsEng.exe", 0&
            .Add "<SysRoot>\system32\Windows Media\Server\nsum.exe", 0&
            .Add "<SysRoot>\system32\MSTask.exe", 0&
            .Add "<SysRoot>\system32\tcpsvcs.exe", 0&
            .Add "<SysRoot>\system32\inetsrv\inetinfo.exe", 0&
            .Add "<SysRoot>\system32\sfmprint.exe", 0&
            .Add "<SysRoot>\System32\snmp.exe", 0&
            .Add "<SysRoot>\System32\ias.dll", 0&
            .Add "<SysRoot>\system32\Windows Media\Server\nspm.exe", 0&
            .Add "<SysRoot>\system32\Windows Media\Server\nspmon.exe", 0&
            .Add "<SysRoot>\system32\Windows Media\Server\nscm.exe", 0&
            .Add "<SysRoot>\system32\regsvc.exe", 0&
            .Add "<SysRoot>\System32\llssrv.exe", 0&
            .Add "<SysRoot>\System32\termsrv.exe", 0&
            .Add "<SysRoot>\system32\RsFsa.exe", 0&
            .Add "<SysRoot>\system32\sfmsvc.exe", 0&
            .Add "<SysRoot>\system32\tlntsvr.exe", 0&
            .Add "<SysRoot>\system32\grovel.exe", 0&
            .Add "<SysRoot>\system32\netdde.exe", 0&
            .Add "<SysRoot>\System32\UtilMan.exe", 0&
            .Add "<SysRoot>\system32\clipsrv.exe", 0&
            .Add "<SysRoot>\system32\Windows Media\NSLite\nslservice.exe", 0&
            .Add "<SysRoot>\system32\faxsvc.exe", 0&
            .Add "<SysRoot>\system32\tftpd.exe", 0&
            .Add "<SysRoot>\SysWow64\mnmsrvc.exe", 0&
            .Add "<SysRoot>\WindowsMobile\wcescomm.dll", 0&
            .Add "<SysRoot>\WindowsMobile\rapimgr.dll", 0&
            
            'Windows Defender
            .Add "<PF64>\Windows Defender\mpsvc.dll", 0&
            .Add "<PF64>\Windows Defender\NisSrv.exe", 0&
            .Add "<PF64>\Windows Defender\MsMpEng.exe", 0&
            .Add "<PF64>\Microsoft Security Client\MsMpEng.exe", 0&
            .Add "<PF64>\Microsoft Security Client\NisSrv.exe", 0&
            .Add "<PF64>\Windows Defender Advanced Threat Protection\MsSense.exe", 0&
            
            For Each vKey In .Keys
                prefix = Left$(vKey, InStr(vKey, "\") - 1)
                Select Case prefix
                    Case "<SysRoot>"
                        .Add Replace$(vKey, prefix, sWinDir), 0&
                    Case "<PF64>"
                        .Add Replace$(vKey, prefix, PF_64), 0&
                    Case "<PF32>"
                        If OSver.IsWin64 Then
                            .Add Replace$(vKey, prefix, PF_32), 0&
                        End If
                End Select
            Next
        End With
    End If
    
    If oDictSRV.Exists(sFilePath) Then
        sArgBase = oDictSRV(sFilePath)
        If sArgBase = 0 Then
            'no arguments defined in database
            IsWinServiceFileName = True
        Else
            'check also an argument
            If StrComp(sArgument, sArgBase, 1) = 0 Then IsWinServiceFileName = True
        End If
    End If
    
    If Not IsWinServiceFileName Then
        'by filename
        Dim colFN As Collection
        Set colFN = New Collection
        
        colFN.Add "aspnet_state.exe"
        'random folder name
        'C:\ProgramData\Microsoft\Windows Defender\platform\4.18.1806.18062-0\MsMpEng.exe
        'C:\ProgramData\Microsoft\Windows Defender\platform\4.18.1806.18062-0\NisSrv.exe
        colFN.Add "MsMpEng.exe"
        colFN.Add "NisSrv.exe"
        colFN.Add "MpCmdRun.exe" 'task
        'C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\WPF\WPFFontCache_v0400.exe
        colFN.Add "WPFFontCache_v0400.exe" 'XP
        colFN.Add "elevation_service.exe" 'Microsoft Edge
        
        sFilename = GetFileNameAndExt(sFilePath)
        
        If isCollectionItemExists(sFilename, colFN) Then IsWinServiceFileName = True
        Set colFN = Nothing
    End If
    
    'if service file is not in list, check if it protected by SFC, excepting AV / Firewall services
    'also, separate blacklist nedeed to identify dangerous host-files like cmd.exe / powershell e.t.c.
    
    If Not IsWinServiceFileName Then
    
'        If Not (StrComp(sFilePath, PF_64 & "\Windows Defender\mpsvc.dll", 1) = 0) _
'          And Not (StrComp(sFilePath, PF_64 & "\Windows Defender\NisSrv.exe", 1) = 0) _
'          And Not (StrComp(sFilePath, PF_64 & "\Windows Defender\MsMpEng.exe", 1) = 0) _
'          And Not (StrComp(sFilePath, PF_64 & "\Microsoft Security Client\MsMpEng.exe", 1) = 0) _
'          And Not (StrComp(sFilePath, PF_64 & "\Microsoft Security Client\NisSrv.exe", 1) = 0) _
'          And Not (StrComp(sFilePath, PF_64 & "\Windows Defender Advanced Threat Protection\MsSense.exe", 1) = 0) Then

'            If Not IsSecurityProductName(sServiceName) Then
'
'                sCompany = GetFilePropCompany(sFilePath)
'                If InStr(1, sCompany, "Microsoft", 1) > 0 Or InStr(1, sCompany, "Корпорация Майкрософт", 1) > 0 Then
'                    IsWinServiceFileName = True
'                End If
'            End If
            
            If IsFileSFC(sFilePath) Then
                
                sFilename = GetFileName(sFilePath, True)
                
                If Not inArraySerialized(sFilename, "rundll32.exe|schtasks.exe|sc.exe|cmd.exe|wscript.exe|" & _
                                                    "mshta.exe|pcalua.exe|powershell.exe", "|", , , vbTextCompare) Then
                    IsWinServiceFileName = True
                End If
            End If

        'End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "IsWinServiceFileName", "File: " & sFilePath
    If inIDE Then Stop: Resume Next
End Function

'Function IsSecurityProductName(sProductName As String) As Boolean
'    Static isInit As Boolean
'    Static AV() As String
'    Dim i&
'
'    If Not isInit Then
'        isInit = True
'        AddToArray AV, "security"
'        AddToArray AV, "antivirus"
'        AddToArray AV, "firewall"
'        AddToArray AV, "protect"
'        AddToArray AV, "Ad-aware"
'        AddToArray AV, "Avast"
'        AddToArray AV, "AVG"
'        AddToArray AV, "Avira"
'        AddToArray AV, "Baidu"
'        AddToArray AV, "BitDefender"
'        AddToArray AV, "Comodo"
'        AddToArray AV, "DrWeb"
'        AddToArray AV, "Emsisoft"
'        AddToArray AV, "ESET"
'        AddToArray AV, "F-Secure"
'        AddToArray AV, "GData"
'        AddToArray AV, "Hitman"
'        AddToArray AV, "Kaspersky"
'        AddToArray AV, "Malwarebytes"
'        AddToArray AV, "McAfee"
'        AddToArray AV, "Norton"
'        AddToArray AV, "Panda"
'        AddToArray AV, "Qihoo"
'        AddToArray AV, "Symantec"
'        AddToArray AV, "TrendMicro"
'        AddToArray AV, "Vipre"
'        AddToArray AV, "Zillya"
'        AddToArray AV, "360"
'    End If
'
'    For i = 0 To UBound(AV)
'        If InStr(1, sProductName, AV(i), 1) <> 0 Then IsSecurityProductName = True: Exit Function
'    Next
'End Function

Public Sub FixO23Item(sItem$, result As SCAN_RESULT)
    'stop & disable & delete NT service
    'O23 - Service: <displayname> - <company> - <file>
    ' (file missing) or (filesize .., MD5 ..) can be appended
    If Not bIsWinNT Then Exit Sub
    
    On Error GoTo ErrorHandler:
    FixIt result
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
    Dim sSource$, sSubscr$, sName$, sHit$, Wow64key As Boolean, result As SCAN_RESULT
    
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
            If (Not (LCase$(sSource) = "about:home" And LCase$(sSubscr) = "about:home") And _
               Not (UCase$(sSource) = "131A6951-7F78-11D0-A979-00C04FD705A2" And UCase$(sSubscr) = "131A6951-7F78-11D0-A979-00C04FD705A2")) _
               Or Not bHideMicrosoft Then
                
                'Example: <Windows folder>\screen.html
                sSource = Replace$(sSource, "<Windows folder>", sWinDir, , , 1)
                sSource = Replace$(sSource, "<System>", sWinSysDir, , , 1)
                If Left$(sSource, 8) = "file:///" Then sSource = Mid$(sSource, 9)
                
                'If file system object
                If Mid$(sSource, 2, 1) = ":" Then
                    sSource = FormatFileMissing(sSource)
                End If
                
                sHit = "O24 - Desktop Component " & sComponents(i) & ": " & sName & " - " & _
                    IIf(sSource <> "", "[Source] = " & sSource, IIf(sSubscr <> "", "[SubscribedURL] = " & sSubscr, "(no file)"))
                
                If Not IsOnIgnoreList(sHit) Then
                    With result
                        .Alias = "O24"
                        .Section = "O24"
                        .HitLineW = sHit
                        AddRegToFix .Reg, REMOVE_KEY, HKEY_CURRENT_USER, sDCKey & "\" & sComponents(i), , , CLng(Wow64key)
                        If sSource <> "" Then AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sSource
                        If sSubscr <> "" Then AddFileToFix .File, REMOVE_FILE Or USE_FEATURE_DISABLE, sSubscr
                        .CureType = REGISTRY_BASED Or FILE_BASED
                    End With
                    AddToScanResults result
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

Public Sub FixO24Item(sItem$, result As SCAN_RESULT)
    On Error GoTo ErrorHandler:
    'delete the entire registry key
    'O24 - Desktop Component 1: Internet Explorer Channel Bar - 131A6951-7F78-11D0-A979-00C04FD705A2
    'O24 - Desktop Component 2: Security - %windir%\index.html
    FixRegistryHandler result
    FixFileHandler result
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_FixO23Item", "sItem=", sItem
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixO24Item_Post()
    Const SPIF_UPDATEINIFILE As Long = 1&
    
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, 0&, SPIF_UPDATEINIFILE 'SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    SleepNoLock 1000
    RestartExplorer
End Sub

Public Sub RestartExplorer()
    'We could do it in official way, e.g. with RestartManager: https://jiangsheng.net/2013/01/22/how-to-restart-windows-explorer-programmatically-using-restart-manager/
    'but, consider we are dealing with malware, it is better to just kill process without notifying loaded modules about this action
    
    ShutdownExplorer
    SleepNoLock 1000
    
    ' Run unelevated (downgrade privileges)
    ' Same as CreateExplorerShellUnelevatedTask task, that uses /NOUACCHECK switch to override task policy
    ' I guess that switch used in task scheduler to prevent recurse call
    Proc.ProcessRunUnelevated2 sWinDir & "\" & "explorer.exe"
End Sub

Public Sub ShutdownExplorer()
    KillProcessByFile sWinDir & "\" & "explorer.exe", True
End Sub
    
Public Function IsOnIgnoreList(sHit$, Optional UpdateList As Boolean, Optional EraseList As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "IsOnIgnoreList - Begin", "Line: " & sHit
    
    Static isInit As Boolean
    Static aIgnoreList() As String
    
    If EraseList Then
        ReDim aIgnoreList(0)
        Exit Function
    End If
    
    If isInit And Not UpdateList Then
        If inArray(sHit, aIgnoreList) Then IsOnIgnoreList = True
    Else
        Dim iIgnoreNum&, i&
        
        isInit = True
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
    Dim sMsg$, sParameters$, hResult$, HRESULT_LastDll$, sErrDesc$, iErrNum&, iErrLastDll&, i&
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
        hResult = ErrMessageText(CLng(iErrNum))
    End If
    
    If iErrLastDll <> 0 Then
        HRESULT_LastDll = ErrMessageText(iErrLastDll)
    End If
    
    For i = 0 To UBound(vCodeModule)
        If Not IsMissing(vCodeModule(i)) Then
            sParameters = sParameters & vCodeModule(i) & " "
        End If
    Next
    
    If AryItems(TranslateNative) Then
        sErrHeader = TranslateNative(590)
    End If
    If 0 = Len(sErrHeader) Then
        If AryItems(Translate) Then
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
        OSData = OSver.Bitness & " " & OSver.OSName & IIf(OSver.Edition <> "", " (" & OSver.Edition & ")", "") & ", " & _
            OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & ", " & _
            "Service Pack: " & OSver.SPVer & "" & IIf(OSver.IsSafeBoot, " (Safe Boot)", "")
    End If
    
    sMsg = sErrHeader & " " & _
        sProcedure & vbCrLf & _
        "Error # " & iErrNum & IIf(iErrNum <> 0, " - " & sErrDesc, "") & _
        vbCrLf & "HRESULT: " & hResult & _
        vbCrLf & "LastDllError # " & iErrLastDll & IIf(iErrLastDll <> 0, " (" & HRESULT_LastDll & ")", "") & _
        vbCrLf & "Trace info: " & sParameters & _
        vbCrLf & vbCrLf & "Windows version: " & OSData & _
        vbCrLf & AppVerPlusName & vbCrLf & _
        "--- EOF ---"
    
    '"Windows version: " & sWinVersion & vbCrLf & vbCrLf & AppVer
    
    If Not bAutoLogSilent Then
        'Clipboard.Clear
        'ClipboardSetText sMsg
        ClipboardSetText sMsg
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
    If iErrNum <> 0 Then ErrText = ErrText & " (" & sErrDesc & ")" & IIf(Len(hResult) <> 0, " (" & hResult & ")", "")
    ErrText = ErrText & " LastDllError = " & iErrLastDll
    If iErrLastDll <> 0 Then ErrText = ErrText & " (" & HRESULT_LastDll & ")"
    If Len(sParameters) <> 0 Then ErrText = ErrText & " " & sParameters
    
    Debug.Print ErrText
    
    ErrReport = ErrReport & vbCrLf & _
        "- " & DateTime & ErrText
    
    AppendErrorLogCustom ">>> ERROR:" & vbCrLf & _
        "- " & DateTime & ErrText
    
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

Public Function ClipboardSetText(sText As String) As Boolean
    Dim hMem As Long
    Dim ptr As Long
    If OpenClipboard(0) Then
        If Len(sText) = 0 Then
            ClipboardSetText = EmptyClipboard()
        Else
            hMem = GlobalAlloc(GMEM_MOVEABLE, 4)
            If hMem <> 0 Then
                ptr = GlobalLock(hMem)
                If ptr <> 0 Then
                    GetMem4 OSver.LangNonUnicodeCode, ByVal ptr
                    GlobalUnlock hMem
                    SetClipboardData CF_LOCALE, hMem
                End If
            End If
            hMem = GlobalAlloc(GMEM_MOVEABLE, LenB(sText) + 2)
            If hMem <> 0 Then
                ptr = GlobalLock(hMem)
                If ptr <> 0 Then
                    lstrcpyn ByVal ptr, ByVal StrPtr(sText), LenB(sText)
                    GlobalUnlock hMem
                    ClipboardSetText = SetClipboardData(CF_UNICODETEXT, hMem)
                End If
            End If
        End If
        CloseClipboard
    End If
End Function

Public Sub AppendErrorLogNoErr(ErrObj As ErrObject, sProcedure As String, ParamArray CodeModule())
    'to append error log without displaying error message to user
    
    On Error Resume Next
    
    Dim i           As Long
    Dim DateTime    As String
    Dim ErrText     As String
    Dim sErrDesc    As String
    Dim iErrNum     As Long
    Dim iErrLastDll As Long
    Dim hResult     As String
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
        hResult = ErrMessageText(CLng(iErrNum))
    End If
    
    If iErrLastDll <> 0 Then
        HRESULT_LastDll = ErrMessageText(iErrLastDll)
    End If
    
    For i = 0 To UBound(CodeModule)
        sParameters = sParameters & CodeModule(i) & " "
    Next

    ErrText = " - " & sProcedure & " - #" & iErrNum
    If iErrNum <> 0 Then ErrText = ErrText & " (" & sErrDesc & ")" & IIf(Len(hResult) <> 0, " (" & hResult & ")", "")
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
    Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800&
    
    Dim sRtrnMsg   As String
    Dim lret       As Long
    Dim hLib       As Long
    
    sRtrnMsg = Space$(MAX_PATH)
    hLib = GetModuleHandle(StrPtr("wininet.dll"))
    If hLib = 0 Then
        hLib = LoadLibrary(StrPtr("wininet.dll"))
    End If
    lret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_FROM_HMODULE Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal hLib, lCode, 0&, StrPtr(sRtrnMsg), MAX_PATH, 0&)
    If lret > 0 Then
        ErrMessageText = Left$(sRtrnMsg, lret)
        ErrMessageText = Replace$(ErrMessageText, vbCrLf, vbNullString)
    End If
End Function

'Public Sub CheckDateFormat()
'    Dim sBuffer$, uST As SYSTEMTIME
'    With uST
'        .wDay = 10
'        .wMonth = 11
'        .wYear = 2003
'    End With
'    sBuffer = String$(255, 0)
'    GetDateFormat 0&, 0&, uST, 0&, StrPtr(sBuffer), 255&
'    sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
'
'    'last try with GetLocaleInfo didn't work on Win2k/XP
'    If InStr(sBuffer, "10") < InStr(sBuffer, "11") Then
'        bIsUSADateFormat = False
'        'msgboxW "sBuffer = " & sBuffer & vbCrLf & "10 < 11, so bIsUSADateFormat False"
'    Else
'        bIsUSADateFormat = True
'        'msgboxW sBuffer & vbCrLf & "10 !< 11, so bIsUSADateFormat True"
'    End If
'
'    'Dim lLndID&, sDateFormat$
'    'lLndID = GetSystemDefaultLCID()
'    'sDateFormat = String$(255, 0)
'    'GetLocaleInfo lLndID, LOCALE_SSHORTDATE, sDateFormat, 255
'    'sDateFormat = left$(sDateFormat, InStr(sDateFormat, vbnullchar) - 1)
'    'If sDateFormat = vbNullString Then Exit Sub
'    ''sDateFormat = "dd-MM-yy" or "M/d/yy"
'    ''I hope this works - dunno what happens in
'    ''yyyy-mm-dd or yyyy-dd-mm format
'    'If InStr(1, sDateFormat, "d", vbTextCompare) < _
'    '   InStr(1, sDateFormat, "m", vbTextCompare) Then
'    '    bIsUSADateFormat = False
'    'Else
'    '    bIsUSADateFormat = True
'    'End If
'End Sub

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

Public Function CheckForReadOnlyMedia() As Boolean
    Dim sMsg$, hFile As Long, sTempFile$, hTransaction&
    
    AppendErrorLogCustom "CheckForReadOnlyMedia - Begin"
    
'    sTempFile = BuildPath(AppPath(), "~dummy.tmp")
'
''    If OSver.IsWindowsVistaOrGreater Then
''
''        hTransaction = CreateTransaction(0, 0, 0, 0, 0, 0, StrPtr("HiJackThis_dummy"))
''
''        If hTransaction <> INVALID_HANDLE_VALUE Then
''            hFile = CreateFileTransacted(StrPtr(sTempFile), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, ByVal 0&, hTransaction, 0&, 0&)
''
''            'ERROR_TRANSACTIONAL_CONFLICT Why ???
''
''            CloseHandle hTransaction
''        End If
''    Else
''
''    End If
'
'    hFile = CreateFile(StrPtr(sTempFile), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, ByVal 0&)
'
'    If hFile <= 0 Then

    If Not CheckFileAccess(AppPath(), GENERIC_WRITE) Then
    
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
'        CloseW hFile
        CheckForReadOnlyMedia = True
    End If
    
'    DeleteFileWEx (StrPtr(sTempFile))
    
    AppendErrorLogCustom "CheckForReadOnlyMedia - End"
End Function

Public Sub SetAllFontCharset(frm As Form, Optional sFontName As String, Optional sFontSize As String, Optional bFontBold As Boolean)
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "SetAllFontCharset - Begin"

    Dim Ctl         As Control
    Dim ctlBtn      As CommandButton
    Dim ctlOptBtn   As OptionButton
    Dim ctlCheckBox As CheckBox
    Dim ctlTxtBox   As TextBox
    Dim ctlLstBox   As ListBox
    Dim CtlLbl      As Label
    Dim CtlFrame    As Frame
    Dim CtlCombo    As ComboBox
    Dim CtlTree     As TreeView
    Dim CtlPict     As PictureBox
    
    For Each Ctl In frm.Controls
        Select Case TypeName(Ctl)
            Case "CommandButton"
                Set ctlBtn = Ctl
                SetFontCharSet ctlBtn, sFontName, sFontSize, bFontBold
            Case "OptionButton"
                Set ctlOptBtn = Ctl
                SetFontCharSet ctlOptBtn, sFontName, sFontSize, bFontBold
            Case "TextBox"
                Set ctlTxtBox = Ctl
                SetFontCharSet ctlTxtBox, sFontName, sFontSize, bFontBold
            Case "ListBox"
                Set ctlLstBox = Ctl
                SetFontCharSet ctlLstBox, sFontName, sFontSize, bFontBold
            Case "Label"
                Set CtlLbl = Ctl
                SetFontCharSet CtlLbl, sFontName, sFontSize, bFontBold
            Case "CheckBox"
                Set ctlCheckBox = Ctl
                'If ctlCheckBox.Name <> "chkConfigTabs" Then
                    SetFontCharSet ctlCheckBox, sFontName, sFontSize, bFontBold
                'End If
            Case "Frame"
                Set CtlFrame = Ctl
                SetFontCharSet CtlFrame, sFontName, sFontSize, bFontBold
            Case "ComboBox"
                Set CtlCombo = Ctl
                If CtlCombo.Name <> "cmbFont" And CtlCombo.Name <> "cmbFontSize" Then
                    SetFontCharSet CtlCombo, sFontName, sFontSize, bFontBold
                End If
            Case "TreeView"
                Set CtlTree = Ctl
                SetFontCharSet CtlTree, sFontName, sFontSize, bFontBold
            Case "PictureBox"
                Set CtlPict = Ctl
                SetFontCharSet CtlPict, sFontName, sFontSize, bFontBold
        End Select
    Next Ctl
    
    AppendErrorLogCustom "SetAllFontCharset - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_SetAllFontCharset"
    If inIDE Then Stop: Resume Next
End Sub

'reset font to initial defaults
Public Sub SetFontDefaults(Ctl As Control, Optional bRelease As Boolean)

    'Here we are saving default state of control and change the state to defaults before changing font,
    'because previous font can be such that has no some property (like it can be BOLD only).
    'In such case after changing font will be alsways BOLDed.

    If bRelease Then
        Set dFontDefault = Nothing
        Erase aFontDefProp
        Exit Sub
    End If
    
    If (dFontDefault Is Nothing) Then
        Set dFontDefault = New clsTrickHashTable
        ReDim aFontDefProp(0)
    End If
    
    Dim CtlPath As String
    Dim idx As Long
    CtlPath = Ctl.Parent.Name & "." & Ctl.Name
    
    If dFontDefault.Exists(CtlPath) Then
        idx = dFontDefault(CtlPath)
        With Ctl.Font
            .Name = "Tahoma"
            .Charset = DEFAULT_CHARSET
            .Weight = 400
            .Bold = aFontDefProp(idx).Bold
            .Italic = aFontDefProp(idx).Italic
            .Underline = aFontDefProp(idx).Underline
            .Size = aFontDefProp(idx).Size
            .StrikeThrough = False
        End With
    Else
        idx = UBound(aFontDefProp) + 1
        dFontDefault.Add CtlPath, idx
        ReDim Preserve aFontDefProp(idx)
        With aFontDefProp(idx)
            .Bold = Ctl.Font.Bold
            .Italic = Ctl.Font.Italic
            .Underline = Ctl.Font.Underline
            .Size = Ctl.Font.Size
        End With
    End If
End Sub

'return BOOL, whether g_FontOnInterface allow to change the font of supplied control
Private Function IsFontAllowedForControl(Ctl As Control) As Boolean
    Static CtlList() As String
    Dim CtlPath As String
    
    If g_FontOnInterface Then
        IsFontAllowedForControl = True
    Else
        CtlPath = Ctl.Parent.Name & "." & Ctl.Name
        
        If 0 = AryPtr(CtlList) Then
            ReDim CtlList(14)
            CtlList(0) = "frmMain.lstResults"
            CtlList(1) = "frmMain.lstIgnore"
            CtlList(2) = "frmMain.lstBackups"
            CtlList(3) = "frmMain.lstHostsMan"
            CtlList(4) = "frmStartupList2.tvwMain"
            CtlList(5) = "frmADSspy.lstADSFound"
            CtlList(6) = "frmADSspy.txtADSContent"
            CtlList(7) = "frmADSspy.txtScanFolder"
            CtlList(8) = "frmCheckDigiSign.txtPaths"
            CtlList(9) = "frmCheckDigiSign.txtExtensions"
            CtlList(10) = "frmProcMan.lstProcessManager"
            CtlList(11) = "frmProcMan.lstProcManDLLs"
            CtlList(12) = "frmUninstMan.lstUninstMan"
            CtlList(13) = "frmUninstMan.txtName"
            CtlList(14) = "frmUnlockRegKey.txtKeys"
        End If
        
        If inArray(CtlPath, CtlList, , , 1) Then
            IsFontAllowedForControl = True
        End If
    End If
End Function

Public Sub SetFontCharSet(Ctl As Control, Optional sFontName As String, Optional sFontSize As String, Optional bFontBold As Boolean)
    On Error GoTo ErrorHandler:
    
    'A big thanks to 'Gun' and 'Adult', two Japanese users
    'who helped me greatly with this
    
    'https://msdn.microsoft.com/en-us/library/aa241713(v=vs.60).aspx
    
    Static isInit As Boolean
    Static lLCID As Long
    
    Dim bNonUsCharset As Boolean
    Dim ControlFont As Font
    Dim lFontSize As Long
    
    '//TODO:
    'Set default Hewbrew 'Non-Unicode: Hebrew (0x40D)' to Arial Unicode MS (after testing)
    
    SetFontDefaults Ctl
    
    'check g_FontOnInterface condition
    If Not IsFontAllowedForControl(Ctl) Then
        Ctl.Font.Charset = DEFAULT_CHARSET
        Exit Sub
    End If
    
    Set ControlFont = Ctl.Font
    
    If Len(sFontName) <> 0 And sFontName <> "Automatic" Then 'if font specified explicitly by user
        ControlFont.Name = sFontName
        
        If sFontSize = "Auto" Or Len(sFontSize) = 0 Then
            lFontSize = 8
        Else
            lFontSize = CStr(sFontSize)
        End If
        ControlFont.Size = lFontSize
        
        'if Hebrew
        'https://msdn.microsoft.com/en-us/library/cc194829.aspx
        
        If OSver.LangDisplayCode = &H40D& Or OSver.LangNonUnicodeCode = &H40D& Then
            ControlFont.Charset = HEBREW_CHARSET
        End If
        ControlFont.Bold = bFontBold
        
        Exit Sub
    End If
    
    bNonUsCharset = True
    
    If Not isInit Then
        lLCID = GetUserDefaultLCID()
        isInit = True
    End If
    
    'Hebrew default behaviour -> choose "Miriam", Size 10 (thanks to @limelect for tests)
    If OSver.LangDisplayCode = &H40D& Or OSver.LangNonUnicodeCode = &H40D& Then
        If FontExist("Miriam") Then
            ControlFont.Name = "Miriam"
            ControlFont.Size = 10
            ControlFont.Charset = HEBREW_CHARSET
            ControlFont.Bold = bFontBold
            Exit Sub
        End If
    End If
    
    Select Case lLCID
         Case &H404 ' Traditional Chinese
            ControlFont.Charset = CHINESEBIG5_CHARSET
            ControlFont.Name = ChrW$(&H65B0) & ChrW$(&H7D30) & ChrW$(&H660E) & ChrW$(&H9AD4)   'New Ming-Li
         Case &H411 ' Japan
            ControlFont.Charset = SHIFTJIS_CHARSET
            ControlFont.Name = ChrW$(&HFF2D) & ChrW$(&HFF33) & ChrW$(&H20) & ChrW$(&HFF30) & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
         Case &H412 ' Korea UserLCID
            ControlFont.Charset = HANGEUL_CHARSET
            ControlFont.Name = ChrW$(&HAD74) & ChrW$(&HB9BC)
         Case &H804 ' Simplified Chinese
            ControlFont.Charset = CHINESESIMPLIFIED_CHARSET
            ControlFont.Name = ChrW$(&H5B8B) & ChrW$(&H4F53)
         Case Else ' The other countries
            ControlFont.Charset = DEFAULT_CHARSET
            ControlFont.Name = "Tahoma"
            bNonUsCharset = False
    End Select
    
    If sFontSize = "Auto" Or Len(sFontSize) = 0 Then
        If bNonUsCharset Then
            lFontSize = 9
        Else
            lFontSize = 8
        End If
    Else
        lFontSize = CLng(sFontSize)
    End If
    
    ControlFont.Size = lFontSize
    ControlFont.Bold = bFontBold
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modMain_SetFontCharSet"
    If inIDE Then Stop: Resume Next
End Sub

Public Function FontExist(sFontName As String) As Boolean
    Dim i As Long
    For i = 0 To Screen.FontCount - 1
        If StrComp(sFontName, Screen.Fonts(i), 1) = 0 Then
            FontExist = True
            Exit For
        End If
    Next i
End Function

Private Function TrimNull(s$) As String
    TrimNull = Left$(s, lstrlen(StrPtr(s)))
End Function

Public Function CheckForStartedFromTempDir() As Boolean
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
        
        If RunFromTemp And (g_sCommandLine = "") Then
            'msgboxW "Запуск из архива запрещен !" & vbCrLf & "Распаковать на рабочий стол для Вас ?", vbExclamation, AppName
            If MsgBoxW(sMsg, vbExclamation Or vbYesNo, g_AppName) = vbYes Then
                Dim NewFile As String
                NewFile = Desktop & "\HiJackThis\" & AppExeName(True)
                MkDirW NewFile, True
                If FileExists(NewFile) Then     ', Cache:=NO_CACHE
                    SetFileAttributes StrPtr(NewFile), GetFileAttributes(StrPtr(NewFile)) And Not FILE_ATTRIBUTE_READONLY
                    DeleteFileWEx StrPtr(NewFile)
                End If
                CopyFile StrPtr(AppPath(True)), StrPtr(NewFile), ByVal 0&
                If FileExists(NewFile) Then     ', Cache:=NO_CACHE
                    frmMain.ReleaseMutex
                    Proc.ProcessRun NewFile     ', "/twice"
                    CheckForStartedFromTempDir = True
                Else
                    'Could not unzip file to Desktop! Please, unzip it manually.
                    MsgBoxW Translate(1007), vbCritical
                    CheckForStartedFromTempDir = True
                    End
                End If
            Else
                CheckForStartedFromTempDir = True
            End If
        End If
    End If
    
    AppendErrorLogCustom "CheckForStartedFromTempDir - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CheckForStartedFromTempDir"
    If inIDE Then Stop: Resume Next
End Function

Public Sub RestartSystem(Optional sExtraPrompt$, Optional bSilent As Boolean, Optional bForceRestartOnServer As Boolean)
    Dim OpSysSet As Object
    Dim OpSys As Object
    Dim lret As Long
    
    If OSver.IsServer Then
        'The server needs to be rebooted to complete required operations. Please, do it on your own.
        MsgBoxW TranslateNative(352), vbInformation
        Exit Sub
    End If
    
    'HiJackThis needs to restart the system to apply the changes.
    'Please, save your work and press 'YES' if you agree to reboot now.
    If Not bSilent Then
        If MsgBoxW(IIf(Len(sExtraPrompt) <> 0, sExtraPrompt & vbCrLf & vbCrLf, "") & TranslateNative(350), vbYesNo Or vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    SetCurrentProcessPrivileges "SeRemoteShutdownPrivilege"
    
    If bIsWinNT Then
        'SHRestartSystemMB g_HwndMain, StrConv(sExtraPrompt & IIf(sExtraPrompt <> vbNullString, vbCrLf & vbCrLf, vbNullString), vbUnicode), 2
        
        If OSver.IsWindowsVistaOrGreater Then
            lret = ExitWindowsEx(EWX_REBOOT Or EWX_FORCEIFHUNG, SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION Or SHTDN_REASON_FLAG_PLANNED)
            'lRet = ExitWindowsEx(EWX_RESTARTAPPS Or EWX_FORCEIFHUNG, SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION Or SHTDN_REASON_FLAG_PLANNED)
        Else 'XP/2000
            lret = ExitWindowsEx(EWX_REBOOT Or EWX_FORCEIFHUNG, SHTDN_REASON_MAJOR_APPLICATION Or SHTDN_REASON_MINOR_INSTALLATION Or SHTDN_REASON_FLAG_PLANNED)
        End If
        
        If lret = 0 Then 'if ExitWindowsEx somehow failed
            Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
            For Each OpSys In OpSysSet
                OpSys.Reboot
            Next
        End If
        
    Else
        SHRestartSystemMB g_HwndMain, sExtraPrompt, 0
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
    
    If AryItems(bufSid) Then
    
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
    Dim sDummy$, sFilename$
    'sCmd is all command-line parameters, like this
    '/param1 /deleteonreboot c:\progra~1\bla\bla.exe /param3
    '/param1 /deleteonreboot "c:\program files\bla\bla.exe" /param3
    
    '/deleteonreboot
    sDummy = Mid$(sCmd, InStr(sCmd, "deleteonreboot") + Len("deleteonreboot") + 1)
    If InStr(sDummy, """") = 1 Then
        'enclosed in quotes, chop off at next quote
        sFilename = Mid$(sDummy, 2)
        sFilename = Left$(sFilename, InStr(sFilename, """") - 1)
    Else
        'no quotes, chop off at next space if present
        If InStr(sDummy, " ") > 0 Then
            sFilename = Left$(sDummy, InStr(sDummy, " ") - 1)
        Else
            sFilename = sDummy
        End If
    End If
    DeleteFileOnReboot sFilename, True
End Sub

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

Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String = " ") As VbMsgBoxResult
    Dim hActiveWnd As Long, hMyWnd As Long, frm As Form
    If inIDE Then
        MsgBoxW = VBA.MsgBox(Prompt, Buttons, Title) 'subclassing walkaround
    Else
        hActiveWnd = GetForegroundWindow()
        For Each frm In Forms
            If frm.hwnd = hActiveWnd Then hMyWnd = hActiveWnd: Exit For
        Next
        MsgBoxW = MessageBox(IIf(hMyWnd <> 0, hMyWnd, g_HwndMain), StrPtr(Prompt), StrPtr(Title), ByVal Buttons)
    End If
End Function

Public Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String = " ") As VbMsgBoxResult
    MsgBox = MsgBoxW(Prompt, Buttons, Title)
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
    
    'Const CSIDL_LOCAL_APPDATA       As Long = &H1C&
    'Const CSIDL_COMMON_PROGRAMS     As Long = &H17&
    'Const FOLDERID_ComputerFolderStr As String = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"

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
    
    'Special note:
    'Under Local System account some environment variables looks like this:
    '
    'APPDATA=C:\Windows\system32\config\systemprofile\AppData\Roaming
    'LOCALAPPDATA=C:\Windows\system32\config\systemprofile\AppData\Local
    'USERPROFILE=C:\Windows\system32\config\systemprofile
    '
    'also these internal variables will be modified:
    '
    'TempCU
    'UserProfile
    'Desktop
    'StartMenuPrograms
    '
    
    AppendErrorLogCustom "InitVariables - Begin"
    
    'Const CSIDL_DESKTOP = 0&
    
    CRCinit
    
    'Init user type array of scan results
    ReInitScanResults
    
    Dim lr As Long, i As Long, nChars As Long
    Dim Path As String, dwBufSize As Long
    
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
    sWinSysDir = sWinDir & "\" & IIf(bIsWinNT, "System32", "System")
    sSysDir = sWinSysDir
    sWinSysDirWow64 = sWinDir & "\SysWOW64"
    
    If bIsWin64 And FolderExists(sWinDir & "\sysnative") And OSver.MajorMinor >= 6 Then
        sSysNativeDir = sWinDir & "\SysNative"
    Else
        sSysNativeDir = sWinDir & "\System32"
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
    
    If GetProfilesDirectory(StrPtr(Path), dwBufSize) Then
        Path = Left$(Path, lstrlen(StrPtr(Path)))
    Else
        If OSver.IsLocalSystemContext Then
            If OSver.IsWindowsVistaOrGreater Then
                Path = SysDisk & "\Users"
            Else
                Path = SysDisk & "\Documents and Settings"
            End If
        Else
            Path = GetParentDir(UserProfile)
        End If
    End If
    ProfilesDir = Path
    
    nChars = MAX_PATH
    AllUsersProfile = String$(nChars, 0)
    If GetAllUsersProfileDirectory(StrPtr(AllUsersProfile), nChars) Then
        AllUsersProfile = Left$(AllUsersProfile, nChars - 1)
    Else
        If Not OSver.IsWindowsVistaOrGreater Then
            If Len(ProfilesDir) = 0 Then
                ProfilesDir = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory")
            End If
            If Len(ProfilesDir) <> 0 Then
                AllUsersProfile = ProfilesDir & "\" & Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "AllUsersProfile")
            End If
        Else    'Win Vista +
            AllUsersProfile = Reg.GetString(HKLM, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProgramData")
        End If
    End If
    If Len(AllUsersProfile) = 0 Then
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
    
    envCurUser = OSver.UserName
    'envCurUser = EnvironW("%UserName%")
    
    ProgramData = EnvironW("%ProgramData%")
    
    'Override some special folders and substitute first found user if token = Local System
    
    If OSver.IsLocalSystemContext Then
    
        Dim ProfileListKey      As String
        Dim ProfileSubKey()     As String
        Dim sSID                As String
    
        ProfileListKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        'ProfilesDirectory = Reg.GetString(HKLM, ProfileListKey, "ProfilesDirectory")
    
        If Reg.EnumSubKeysToArray(HKLM, ProfileListKey, ProfileSubKey()) > 0 Then
            For i = LBound(ProfileSubKey) To UBound(ProfileSubKey)
            
                sSID = ProfileSubKey(i)
            
                If Not (sSID = "S-1-5-18" Or _
                        sSID = "S-1-5-19" Or _
                        sSID = "S-1-5-20") Then
                    
                    UserProfile = Reg.GetString(HKLM, ProfileListKey & "\" & sSID, "ProfileImagePath")
                    
                    'just in case
                    If UserProfile = "" Then UserProfile = SysDisk & "\All Users"
                    
                    AppData = ""
                    LocalAppData = ""
                    Desktop = ""
                    TempCU = ""
                    
                    'если профиль загружен
                    If Reg.KeyExists(HKEY_USERS, sSID) Then
                        
                        AppData = Reg.GetString(HKEY_USERS, sSID & _
                            "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AppData")
                    
                        LocalAppData = Reg.GetString(HKEY_USERS, sSID & _
                            "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Local AppData")
                    
                        Desktop = Reg.GetString(HKEY_USERS, sSID & _
                            "\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Desktop")
                    
                        TempCU = Reg.GetString(HKEY_USERS, sSID & _
                            "\Environment", "TEMP")
                        
                        'HKU contains paths of REG_EXPAND_SZ type with %UserProfile% value, so they will be expanded automatically with wrong values,
                        'so we need to substitute correct values manually:
                        AppData = PathSubstituteProfile(AppData, UserProfile)
                        LocalAppData = PathSubstituteProfile(LocalAppData, UserProfile)
                        Desktop = PathSubstituteProfile(Desktop, UserProfile)
                        TempCU = PathSubstituteProfile(TempCU, UserProfile)
                        
                        If OSver.MajorMinor < 6 Then
                            AppDataLocalLow = AppData
                        Else
                            AppDataLocalLow = BuildPath(GetParentDir(AppData), "LocalLow")
                        End If
                    End If
    
                    'если профиль не загружен
                    
                    If OSver.IsWindowsVistaOrGreater Then
                        
                        If AppData = "" Then AppData = UserProfile & "\AppData\Roaming"
                        If AppDataLocalLow = "" Then AppDataLocalLow = BuildPath(GetParentDir(AppData), "LocalLow")
                        If LocalAppData = "" Then LocalAppData = UserProfile & "\AppData\Local"
                        If Desktop = "" Then Desktop = UserProfile & "\Desktop"
                        If TempCU = "" Then TempCU = LocalAppData & "\Temp"
                        
                    Else
                        If AppData = "" Then AppData = UserProfile & "\Application Data"
                        If AppDataLocalLow = "" Then AppDataLocalLow = AppData
                        If LocalAppData = "" Then LocalAppData = UserProfile & "\Local Settings"
                        If TempCU = "" Then TempCU = LocalAppData & "\Temp"
                        
                        If Desktop = "" Then
                            Path = Reg.GetString(HKLM, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Common Desktop")
                            
                            If Path = "" Then
                                If IsSlavianCultureCode(OSver.LangSystemCode) Then
                                    Desktop = UserProfile & "\" & LoadResString(606) 'Рабочий стол
                                Else
                                    Desktop = UserProfile & "\Desktop"
                                End If
                            Else
                                Path = GetFileNameAndExt(Path)
                                Desktop = UserProfile & "\" & Path
                            End If
                        End If
                    End If
                    
                    Exit For
                
                End If
            Next
        End If
    End If
    
    ' Shortcut interfaces initialization
    'IURL_Init
    ISL_Init
    
    Set oDict.TaskWL_ID = New clsTrickHashTable
    oDict.TaskWL_ID.CompareMode = vbTextCompare
    
    Set colProfiles = New Collection
    GetProfiles
    
    FillUsers
    
    Set cMath = New clsMath
    'Set oRegexp = New cRegExp
    
    LIST_BACKUP_FILE = BuildPath(AppPath(), "Backups\List.ini")
    
    InitBackupIni
    
    If OSver.MajorMinor >= 6.1 Then
        Set TaskBar = New TaskbarList
    End If
    
    Dim HE As clsHiveEnum
    Set HE = New clsHiveEnum
    
    'Call CLSIDFromString(StrPtr(FOLDERID_ComputerFolderStr), FOLDERID_ComputerFolder)
    
    'Load ru phrases
    STR_CONST.RU_LINKS = LoadResString(600)
    STR_CONST.RU_NO = LoadResString(601)
    STR_CONST.UA_CANT_LOAD_LANG = Replace$(LoadResString(602), "\n", vbCrLf)
    STR_CONST.RU_CANT_LOAD_LANG = Replace$(LoadResString(603), "\n", vbCrLf)
    STR_CONST.RU_MICROSOFT = LoadResString(604)
    STR_CONST.RU_PC = LoadResString(605)
    
    AppendErrorLogCustom "InitVariables - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "InitVariables"
    If inIDE Then Stop: Resume Next
End Sub

Public Function PathSubstituteProfile(Path As String, Optional ByVal sUserProfileDir As String) As String
    'Substitute 'sUserProfileDir' to 'Path' if 'Path' goes through %UserProfile%.s
    'Note: sUserProfileDir will be trimmed to c:\Users\User, if it contains more nested directories;
    '      sUserProfileDir can be not a profile at all. In such case substitution is not performed.
    
    Static bInit        As Boolean
    Static sCurUserProfile As String
    
    Dim pos As Long
    
    If Not bInit Then
        bInit = True
        sCurUserProfile = EnvironW("%UserProfile%")
    End If
    
    'specified dir is a profile's dir?
    If StrBeginWith(sUserProfileDir, ProfilesDir & "\") And (Len(sUserProfileDir) > (Len(ProfilesDir) + 1)) Then
        
        'expanded path contains current profile's dir?
        If StrBeginWith(Path, sCurUserProfile & "\") Or StrComp(Path, sCurUserProfile, 1) = 0 Then
            
            'extracting path to profile from the string, if it is specified with additional dirs
            pos = InStr(Len(ProfilesDir) + 2, sUserProfileDir, "\")
            If pos <> 0 Then
                sUserProfileDir = Left$(sUserProfileDir, pos - 1)
            End If
            
            'substitute
            PathSubstituteProfile = BuildPath(sUserProfileDir, Mid$(Path, Len(sCurUserProfile) + 2))
            Exit Function
        End If
    End If
    
    PathSubstituteProfile = Path
End Function

Public Function EnvironW(ByVal SrcEnv As String, Optional UseRedir As Boolean, Optional ByVal sUserProfileDir As String) As String
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
        
        'if need expanding under certain user
        If Len(sUserProfileDir) <> 0 Then
            EnvironW = PathSubstituteProfile(EnvironW, sUserProfileDir)
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
        ReDim SplitSafe(0)
    Else
        SplitSafe = Split(sComplexString, Delimiter)
    End If
End Function

Public Sub ArrayRemoveEmptyItems(arr() As String)
    Dim i As Long
    Dim d As Long
    Dim bShift As Boolean
    
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) <> 0 Then
            If bShift Then arr(d) = arr(i): d = d + 1 'shifting items
        Else
            If Not bShift Then bShift = True: d = i
        End If
    Next
    
    If bShift Then
        If d > LBound(arr) Then
            ReDim Preserve arr(d - 1)
        ElseIf UBound(arr) > LBound(arr) Then
            ReDim Preserve arr(d)
        End If
    End If
End Sub

'get the first item of serilized array
Public Function SplitExGetFirst(sSerializedArray As String, Optional Delimiter As String = " ") As String
    SplitExGetFirst = SplitSafe(sSerializedArray, Delimiter)(0)
End Function

'get the last item of serialized array
Public Function SplitExGetLast(sSerializedArray As String, Optional Delimiter As String = " ") As String
    Dim ret() As String
    ret = SplitSafe(sSerializedArray, Delimiter)
    SplitExGetLast = ret(UBound(ret))
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

'Remove empty strings from array
Public Sub CompressArray(arr() As String)
    On Error GoTo ErrorHandler:

    If 0 = AryPtr(arr) Then Exit Sub
    Dim i As Long
    Dim pIdx As Long
    pIdx = -1
    For i = 0 To UBound(arr)
        If Len(arr(i)) = 0 Then
            If pIdx = -1 Then
                pIdx = i
            End If
        Else
            If pIdx <> -1 Then
                arr(pIdx) = arr(i)
                pIdx = pIdx + 1
            End If
        End If
    Next
    If pIdx > 0 Then
        ReDim Preserve arr(pIdx - 1)
    End If
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CompressArray"
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

Public Function LoadWindowPos(frm As Form, IdSection As SETTINGS_SECTION) As Boolean
    
    If frm.WindowState = vbMinimized Or frm.WindowState = vbMaximized Then Exit Function
    
    LoadWindowPos = True
    
    If IdSection <> SETTINGS_SECTION_MAIN Then
    
        Dim iHeight As Long, iWidth As Long
        iHeight = CLng(RegReadHJT("WinHeight", "-1", , IdSection))
        iWidth = CLng(RegReadHJT("WinWidth", "-1", , IdSection))
        
        If iHeight = -1 Or iWidth = -1 Then LoadWindowPos = False
        
        If iHeight > 0 And iWidth > 0 Then
            If iHeight > Screen.Height Then iHeight = Screen.Height
            If iWidth > Screen.Width Then iWidth = Screen.Width
            
            If iHeight < 500 Then iHeight = 500
            If iWidth < 1000 Then iWidth = 1000
            
            frm.Height = iHeight
            frm.Width = iWidth
        End If
    End If
    
    Dim iTop As Long, iLeft As Long
    iTop = CLng(RegReadHJT("WinTop", "-1", , IdSection))
    iLeft = CLng(RegReadHJT("WinLeft", "-1", , IdSection))
    
    If iTop = -1 Or iLeft = -1 Then
    
        LoadWindowPos = False
        CenterForm frm
    Else
        If iTop > (Screen.Height - 2500) Then iTop = Screen.Height - 2500
        If iLeft > (Screen.Width - 5000) Then iLeft = Screen.Width - 5000
        If iTop < 0 Then iTop = 0
        If iLeft < 0 Then iLeft = 0
        
        frm.Top = iTop
        frm.Left = iLeft
    End If
    
    If CLng(RegReadHJT("WinState", "0", , IdSection)) = vbMaximized Then frm.WindowState = vbMaximized
End Function

Public Sub SaveWindowPos(frm As Form, IdSection As SETTINGS_SECTION)

    If frm.WindowState <> vbMinimized And frm.WindowState <> vbMaximized Then
        RegSaveHJT "WinTop", CStr(frm.Top), IdSection
        RegSaveHJT "WinLeft", CStr(frm.Left), IdSection
        RegSaveHJT "WinHeight", CStr(frm.Height), IdSection
        RegSaveHJT "WinWidth", CStr(frm.Width), IdSection
    End If
    RegSaveHJT "WinState", CStr(frm.WindowState), IdSection
    
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
    
    Dim i&, idx&
    
    If 0 = AryItems(AddArray) Then Exit Sub
    If 0 = AryItems(DestArray) Then
        idx = -1
        ReDim DestArray(UBound(AddArray) - LBound(AddArray))
    Else
        idx = UBound(DestArray)
        ReDim Preserve DestArray(UBound(DestArray) + (UBound(AddArray) - LBound(AddArray)) + 1)
    End If
    
    For i = LBound(AddArray) To UBound(AddArray)
        idx = idx + 1
        DestArray(idx) = AddArray(i)
    Next
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Parser.ConcatArrays"
End Sub

Public Sub QuickSort(j() As String, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler:
    Dim i As Long, L As Long, M As String, wsp As String
    i = low: L = high: M = j((i + L) \ 2)
    Do Until i > L: Do While j(i) < M: i = i + 1: Loop: Do While j(L) > M: L = L - 1: Loop
        If (i <= L) Then wsp = j(i): j(i) = j(L): j(L) = wsp: i = i + 1: L = L - 1
    Loop
    If low < L Then QuickSort j, low, L
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
    'Fixed by Dragokas
    On Error GoTo ErrorHandler:
    Dim lpSTR As Long, ptr As Long, Key As String
    If Col Is Nothing Then Exit Function
    Select Case Index
    Case Is < 1, Is > Col.Count: Exit Function
    Case Else
        ptr = ObjPtr(Col)
        Do While Index
            GetMem4 ByVal ptr + 24, ptr
            Index = Index - 1
        Loop
    End Select
    GetMem4 ByVal VarPtr(Key), lpSTR
    GetMem4 ByVal ptr + 16, ByVal VarPtr(Key)
    GetCollectionKeyByIndex = Key
    GetMem4 lpSTR, ByVal VarPtr(Key)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetCollectionKeyByIndex"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetCollectionIndexByItem(sItem As String, Col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(Col.Item(i), sItem, CompareMode) = 0 Then
            GetCollectionIndexByItem = i
            Exit For
        End If
    Next
End Function

Public Function GetCollectionKeyByItem(sItem As String, Col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As String
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(Col.Item(i), sItem, CompareMode) = 0 Then
            GetCollectionKeyByItem = GetCollectionKeyByIndex(i, Col)
            Exit For
        End If
    Next
End Function

Public Function isCollectionKeyExists(Key As String, Col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(GetCollectionKeyByIndex(i, Col), Key, CompareMode) = 0 Then isCollectionKeyExists = True: Exit For
    Next
End Function

Public Function isCollectionItemExists(sItem As String, Col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As Boolean
    isCollectionItemExists = (GetCollectionIndexByItem(sItem, Col, CompareMode) <> 0)
End Function

Public Function GetCollectionItemByKey(Key As String, Col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As String
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(GetCollectionKeyByIndex(i, Col), Key, CompareMode) = 0 Then GetCollectionItemByKey = Col.Item(i)
    Next
End Function

Public Function GetCollectionIndexByKey(Key As String, Col As Collection, Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(GetCollectionKeyByIndex(i, Col), Key, CompareMode) = 0 Then GetCollectionIndexByKey = i
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

            If AryItems(SubFolders) Then
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
    Dim X As Long, s$
    Dim idx As Long
    
    With lstControl
        For idx = 0 To .ListCount - 1
            s = Replace$(.List(idx), vbTab, "12345678")
            If .Width < frmMain.TextWidth(s) + 300 And X < frmMain.TextWidth(s) + 300 Then
                X = frmMain.TextWidth(.List(idx)) + 300
            End If
        Next
        If X <> 0 Then
            If frmMain.ScaleMode = vbTwips Then X = X / Screen.TwipsPerPixelX + 50  ' if twips change to pixels (+50 to account for the width of the vertical scrollbar
        End If
        SendMessage .hwnd, LB_SETHORIZONTALEXTENT, X, ByVal 0&
    End With
End Sub

'Public Function IsArrDimmed(vArray As Variant) As Boolean
'    IsArrDimmed = (GetArrDims(vArray) > 0)
'End Function

Public Function AryItems(vArray As Variant) As Long
    Dim ppSA As Long
    Dim pSA As Long
    Dim VT As Long
    'Dim sa As SAFEARRAY
    Dim pvData As Long
    Const vbByRef As Integer = 16384

    If IsArray(vArray) Then
        GetMem4 ByVal VarPtr(vArray) + 8, ppSA      ' pV -> ppSA (pSA)
        If ppSA <> 0 Then
            GetMem2 vArray, VT
            If VT And vbByRef Then
                GetMem4 ByVal ppSA, pSA                 ' ppSA -> pSA
            Else
                pSA = ppSA
            End If
            If pSA <> 0 Then
                'memcpy sa, ByVal pSA, LenB(sa)
                'If sa.pvData <> 0 Then
                GetMem4 ByVal pSA + 12, pvData
                If pvData <> 0 Then
                    AryItems = UBound(vArray) - LBound(vArray) + 1
                End If
            End If
        End If
    End If
End Function

Public Function UBoundSafe(vArray As Variant) As Long
    If AryItems(vArray) Then
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
    Static isInit As Boolean
    
    Dim Other       As String
    Dim i           As Long
    For i = 0 To UBound(CodeModule)
        Other = Other & CodeModule(i) & " | "
    Next
    
    Dim tim1 As Currency
    If Not isInit Then
        isInit = True
        QueryPerformanceFrequency freq
    End If
    QueryPerformanceCounter tim1
    
    If bDebugToFile Then
        If g_hDebugLog <> 0 Then
            If InStr(Other, "modFile.PutW") = 0 Then 'prevent infinite loop
                Dim b() As Byte
                b = "- " & time & " - " & Format$(tim1 / freq, "##0.000") & " - " & Other & vbCrLf
                PutW g_hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
            End If
        End If
    End If
    
    If bDebugMode Then
    
        OutputDebugStringA Other

        If Not (ErrLogCustomText Is Nothing) Then
            ErrLogCustomText.Append (vbCrLf & "- " & time & " - " & Format$(tim1 / freq, "##0.000") & " - " & Other)
        End If
    
        'If DebugHeavy Then AddtoLog vbCrLf & "- " & time & " - " & Other
    End If
End Sub

Public Sub OpenDebugLogHandle()
    If g_hDebugLog > 0 Then Exit Sub
    
    If Len(g_sDebugLogFile) = 0 Then
        g_sDebugLogFile = BuildPath(AppPath(), "HiJackThis_debug.log")
    End If
    
    If FileExists(g_sDebugLogFile) Then DeleteFileWEx StrPtr(g_sDebugLogFile), , True
    
    On Error Resume Next
    OpenW g_sDebugLogFile, FOR_OVERWRITE_CREATE, g_hDebugLog, g_FileBackupFlag
    
    If g_hDebugLog <= 0 Then
        g_sDebugLogFile = Left$(g_sDebugLogFile, Len(g_sDebugLogFile) - 4) & "_2.log"
                    
        Call OpenW(g_sDebugLogFile, FOR_OVERWRITE_CREATE, g_hDebugLog)
        
    End If
    
    Dim sCurTime$
    sCurTime = vbCrLf & vbCrLf & "Logging started at: " & Now() & vbCrLf & vbCrLf
    PutW g_hDebugLog, 1&, StrPtr(sCurTime), LenB(sCurTime), doAppend:=True
End Sub

Public Sub OpenLogHandle()
    
    Dim ov As OVERLAPPED
    
    If Len(g_sLogFile) = 0 Then
        g_sLogFile = BuildPath(AppPath(), "HiJackThis.log")
    End If
    
    If FileExists(g_sLogFile, , True) Then DeleteFileWEx StrPtr(g_sLogFile), , True
    
    On Error Resume Next
    OpenW g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog, g_FileBackupFlag
    
    If g_hLog > 0 Then
        ov.offset = 0
        ov.InternalHigh = 0
        ov.hEvent = 0
        
        Dim lret As Long
        
        lret = LockFileEx(g_hLog, LOCKFILE_EXCLUSIVE_LOCK Or LOCKFILE_FAIL_IMMEDIATELY, 0&, 1& * 1024 * 1024, 0&, VarPtr(ov))
        
        If lret Then
            g_LogLocked = True
        Else
            Debug.Print "Can't lock file. Err = " & Err.LastDllError
        End If
    Else
        g_sLogFile = Left$(g_sLogFile, Len(g_sLogFile) - 4) & "_2.log"
        
        Call OpenW(g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog)
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
    If 0 = AryItems(uArray) Then
        ReDim uArray(0)
        uArray(0) = sItem
    Else
        ReDim Preserve uArray(UBound(uArray) + 1)
        uArray(UBound(uArray)) = sItem
    End If
End Sub

Public Sub AddToArrayLong(ByRef uArray As Variant, lItem As Long)
    If 0 = AryItems(uArray) Then
        ReDim uArray(0)
        uArray(0) = lItem
    Else
        ReDim Preserve uArray(UBound(uArray) + 1)
        uArray(UBound(uArray)) = lItem
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

Public Sub LockMenu(bDoUnlock As Boolean)
    With frmMain
        .mnuFile.Enabled = bDoUnlock
        .mnuTools.Enabled = bDoUnlock
        .mnuHelp.Enabled = bDoUnlock
    End With
End Sub

Public Sub LockInterface(bAllowInfoButtons As Boolean, bDoUnlock As Boolean)
    'Lock controls when scanning
    On Error GoTo ErrorHandler:
    
    Dim mnu As Menu
    Dim Ctl As Control
    
    For Each Ctl In frmMain.Controls
        If TypeName(Ctl) = "Menu" Then
            Set mnu = Ctl
            If InStr(1, mnu.Name, "delim", 1) = 0 Then
                mnu.Enabled = bDoUnlock
            End If
        End If
    Next
    Set mnu = Nothing
    Set Ctl = Nothing
    
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

Private Sub GetSpecialFolders_XP_2(sLog As clsStringBuilder)
    On Error GoTo ErrorHandler:
    
    sLog.AppendLine "ADMINTOOLS = " & GetSpecialFolderPath(CSIDL_ADMINTOOLS)
    If OSver.MajorMinor >= 6 Then sLog.AppendLine "ALTSTARTUP = " & GetSpecialFolderPath(CSIDL_ALTSTARTUP) 'Vista+
    sLog.AppendLine "APPDATA = " & GetSpecialFolderPath(CSIDL_APPDATA)
    'sLog.AppendLine "CSIDL_BITBUCKET = " & "(virtual)"
    sLog.AppendLine "CDBURN_AREA = " & GetSpecialFolderPath(CSIDL_CDBURN_AREA)
    sLog.AppendLine "COMMON_ADMINTOOLS = " & GetSpecialFolderPath(CSIDL_COMMON_ADMINTOOLS)
    If OSver.MajorMinor >= 6 Then sLog.AppendLine "COMMON_ALTSTARTUP = " & GetSpecialFolderPath(CSIDL_COMMON_ALTSTARTUP) 'Vista+
    sLog.AppendLine "COMMON_APPDATA = " & GetSpecialFolderPath(CSIDL_COMMON_APPDATA)
    sLog.AppendLine "COMMON_DESKTOPDIRECTORY = " & GetSpecialFolderPath(CSIDL_COMMON_DESKTOPDIRECTORY)
    sLog.AppendLine "COMMON_DOCUMENTS = " & GetSpecialFolderPath(CSIDL_COMMON_DOCUMENTS)
    sLog.AppendLine "COMMON_FAVORITES = " & GetSpecialFolderPath(CSIDL_COMMON_FAVORITES)
    sLog.AppendLine "COMMON_MUSIC = " & GetSpecialFolderPath(CSIDL_COMMON_MUSIC)
    If OSver.MajorMinor >= 6 Then sLog.AppendLine "COMMON_OEM_LINKS = " & GetSpecialFolderPath(CSIDL_COMMON_OEM_LINKS) 'Vista+"
    sLog.AppendLine "COMMON_PICTURES = " & GetSpecialFolderPath(CSIDL_COMMON_PICTURES)
    sLog.AppendLine "COMMON_PROGRAMS = " & GetSpecialFolderPath(CSIDL_COMMON_PROGRAMS)
    sLog.AppendLine "COMMON_STARTMENU = " & GetSpecialFolderPath(CSIDL_COMMON_STARTMENU)
    sLog.AppendLine "COMMON_STARTUP = " & GetSpecialFolderPath(CSIDL_COMMON_STARTUP)
    sLog.AppendLine "COMMON_TEMPLATES = " & GetSpecialFolderPath(CSIDL_COMMON_TEMPLATES)
    sLog.AppendLine "COMMON_VIDEO = " & GetSpecialFolderPath(CSIDL_COMMON_VIDEO)
    'sLog.AppendLine "COMPUTERSNEARME = " & "(virtual)"
    'sLog.AppendLine "CONNECTIONS = " & "(virtual)"
    'sLog.AppendLine "CONTROLS = " & "(virtual)"
    sLog.AppendLine "COOKIES = " & GetSpecialFolderPath(CSIDL_COOKIES)
    sLog.AppendLine "DESKTOP = " & GetSpecialFolderPath(CSIDL_DESKTOP)
    sLog.AppendLine "DESKTOPDIRECTORY = " & GetSpecialFolderPath(CSIDL_DESKTOPDIRECTORY)
    'sLog.AppendLine "DRIVES = " & "(virtual)"
    sLog.AppendLine "FAVORITES = " & GetSpecialFolderPath(CSIDL_FAVORITES)
    sLog.AppendLine "FLAG_CREATE = " & GetSpecialFolderPath(CSIDL_FLAG_CREATE)
    sLog.AppendLine "FLAG_DONT_VERIFY = " & GetSpecialFolderPath(CSIDL_FLAG_DONT_VERIFY)
    sLog.AppendLine "FLAG_MASK = " & GetSpecialFolderPath(CSIDL_FLAG_MASK)
    sLog.AppendLine "FLAG_NO_ALIAS = " & GetSpecialFolderPath(CSIDL_FLAG_NO_ALIAS)
    sLog.AppendLine "FLAG_PER_USER_INIT = " & GetSpecialFolderPath(CSIDL_FLAG_PER_USER_INIT)
    sLog.AppendLine "FONTS = " & GetSpecialFolderPath(CSIDL_FONTS)
    sLog.AppendLine "HISTORY = " & GetSpecialFolderPath(CSIDL_HISTORY)
    'sLog.AppendLine "INTERNET = " & "(virtual)"
    sLog.AppendLine "INTERNET_CACHE = " & GetSpecialFolderPath(CSIDL_INTERNET_CACHE)
    sLog.AppendLine "LOCAL_APPDATA = " & GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)
    'sLog.AppendLine "MYDOCUMENTS = " & "(virtual)"
    sLog.AppendLine "MYMUSIC = " & GetSpecialFolderPath(CSIDL_MYMUSIC)
    sLog.AppendLine "MYPICTURES = " & GetSpecialFolderPath(CSIDL_MYPICTURES)
    sLog.AppendLine "MYVIDEO = " & GetSpecialFolderPath(CSIDL_MYVIDEO)
    sLog.AppendLine "NETHOOD = " & GetSpecialFolderPath(CSIDL_NETHOOD)
    'sLog.AppendLine "NETWORK = " & "(virtual)"
    sLog.AppendLine "PERSONAL = " & GetSpecialFolderPath(CSIDL_PERSONAL)
    'sLog.AppendLine "PRINTERS = " & "(virtual)"
    sLog.AppendLine "PRINTHOOD = " & GetSpecialFolderPath(CSIDL_PRINTHOOD)
    sLog.AppendLine "PROFILE = " & GetSpecialFolderPath(CSIDL_PROFILE)
    sLog.AppendLine "PROGRAM_FILES = " & GetSpecialFolderPath(CSIDL_PROGRAM_FILES)
    sLog.AppendLine "PROGRAM_FILES_COMMON = " & GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMON)
    If OSver.IsWin64 Then
        sLog.AppendLine "PROGRAM_FILES_COMMONX86 = " & GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMONX86) 'x64
        sLog.AppendLine "PROGRAM_FILESX86 = " & GetSpecialFolderPath(CSIDL_PROGRAM_FILESX86) 'x64
    End If
    sLog.AppendLine "PROGRAMS = " & GetSpecialFolderPath(CSIDL_PROGRAMS)
    sLog.AppendLine "RECENT = " & GetSpecialFolderPath(CSIDL_RECENT)
    sLog.AppendLine "RESOURCES = " & GetSpecialFolderPath(CSIDL_RESOURCES)
    sLog.AppendLine "RESOURCES_LOCALIZED = " & GetSpecialFolderPath(CSIDL_RESOURCES_LOCALIZED)
    sLog.AppendLine "SENDTO = " & GetSpecialFolderPath(CSIDL_SENDTO)
    sLog.AppendLine "STARTMENU = " & GetSpecialFolderPath(CSIDL_STARTMENU)
    sLog.AppendLine "STARTUP = " & GetSpecialFolderPath(CSIDL_STARTUP)
    sLog.AppendLine "SYSTEM = " & GetSpecialFolderPath(CSIDL_SYSTEM)
    sLog.AppendLine "SYSTEMX86 = " & GetSpecialFolderPath(CSIDL_SYSTEMX86)
    sLog.AppendLine "TEMPLATES = " & GetSpecialFolderPath(CSIDL_TEMPLATES)
    sLog.AppendLine "WINDOWS = " & GetSpecialFolderPath(CSIDL_WINDOWS)

    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetSpecialFolders_XP_2"
    If inIDE Then Stop: Resume Next
End Sub
    

'Private Sub GetSpecialFolders_XP(sLog As clsStringBuilder)
'    On Error GoTo ErrorHandler:
'
'    '// TODO: Append with GetSpecialFolderPath()
'
'    Dim HE As clsHiveEnum
'    Set HE = New clsHiveEnum
'
'    Dim i As Long
'    Dim aName() As String
'    Dim aValue() As String
'    Dim avValue() As Variant
'    Dim sValue As String
'    Dim Key As Variant
'    Dim dDict As clsTrickHashTable
'    Set dDict = New clsTrickHashTable
'    dDict.CompareMode = vbTextCompare
'
'    HE.Init HE_HIVE_HKCU Or HE_HIVE_HKLM, , HE_REDIR_NO_WOW
'    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
'    HE.AddKey "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
'
'    Do While HE.MoveNext
'        For i = 1 To Reg.EnumValuesToArray(HE.Hive, HE.Key, aValue, HE.Redirected)
'            If InStr(1, aValue(i), "Do not use this registry key", 1) = 0 Then
'                If Not dDict.Exists(aValue(i)) Then
'                    sValue = Reg.GetString(HE.Hive, HE.Key, aValue(i), HE.Redirected)
'                    If Len(sValue) <> 0 Then
'                        dDict.Add aValue(i), sValue
'                    End If
'                End If
'            End If
'        Next
'    Loop
'
'    ReDim aName(dDict.Count - 1)
'    ReDim avValue(dDict.Count - 1)
'    i = 0
'    For Each Key In dDict.Keys
'        aName(i) = Key
'        avValue(i) = dDict(Key)
'        i = i + 1
'    Next
'    QuickSortSpecial aName, avValue, 0, dDict.Count - 1
'    For i = 0 To dDict.Count - 1
'        sLog.AppendLine aName(i) & " = " & avValue(i)
'    Next
'
'    Set dDict = Nothing
'    Set HE = Nothing
'
'    Exit Sub
'ErrorHandler:
'    ErrorMsg Err, "GetSpecialFolders_XP"
'    If inIDE Then Stop: Resume Next
'End Sub

'Private Sub GetSpecialFolders_Vista_2(sLog As clsStringBuilder)
'    On Error GoTo ErrorHandler:
'
'    If OSver.MajorMinor >= 6.2 Then sLog.AppendLine "FOLDERID_AccountPictures = " & GetKnownFolderPath("{008ca0b1-55b4-4c56-b8a8-4de4b299d3be}") '8+
'    'sLog.AppendLine "FOLDERID_AddNewPrograms = " & GetKnownFolderPath_GUID(FOLDERID_AddNewPrograms) 'Virtual
'    sLog.AppendLine "FOLDERID_AdminTools = " & GetKnownFolderPath_GUID(FOLDERID_AdminTools)
'    If OSver.MajorMinor >= 6.2 Then sLog.AppendLine "FOLDERID_ApplicationShortcuts = " & GetKnownFolderPath("{A3918781-E5F2-4890-B3D9-A7E54332328C}") '8+
'    'sLog.AppendLine "FOLDERID_AppsFolder = " & GetKnownFolderPath("{1e87508d-89c2-42f0-8a7e-645a0f50ca58}") 'Virtual
'    'sLog.AppendLine "FOLDERID_AppUpdates = " & GetKnownFolderPath_GUID(FOLDERID_AppUpdates) 'Virtual
'    If OSver.MajorMinor >= 6.3 Then sLog.AppendLine "FOLDERID_CameraRoll = " & GetKnownFolderPath("{AB5FB87B-7CE2-4F83-915D-550846C9537B}") '8.1+
'    sLog.AppendLine "FOLDERID_CDBurning = " & GetKnownFolderPath_GUID(FOLDERID_CDBurning)
'    'sLog.AppendLine "FOLDERID_ChangeRemovePrograms = " & GetKnownFolderPath_GUID(FOLDERID_ChangeRemovePrograms) 'Virtual
'    sLog.AppendLine "FOLDERID_CommonAdminTools = " & GetKnownFolderPath_GUID(FOLDERID_CommonAdminTools)
'    sLog.AppendLine "FOLDERID_CommonOEMLinks = " & GetKnownFolderPath_GUID(FOLDERID_CommonOEMLinks)
'    sLog.AppendLine "FOLDERID_CommonPrograms = " & GetKnownFolderPath_GUID(FOLDERID_CommonPrograms)
'    sLog.AppendLine "FOLDERID_CommonStartMenu = " & GetKnownFolderPath_GUID(FOLDERID_CommonStartMenu)
'    sLog.AppendLine "FOLDERID_CommonStartup = " & GetKnownFolderPath_GUID(FOLDERID_CommonStartup)
'    sLog.AppendLine "FOLDERID_CommonTemplates = " & GetKnownFolderPath_GUID(FOLDERID_CommonTemplates)
'    'sLog.AppendLine "FOLDERID_ComputerFolder = " & GetKnownFolderPath_GUID(FOLDERID_ComputerFolder) 'Virtual
'    'sLog.AppendLine "FOLDERID_ConflictFolder = " & GetKnownFolderPath_GUID(FOLDERID_ConflictFolder) 'Virtual
'    'sLog.AppendLine "FOLDERID_ConnectionsFolder = " & GetKnownFolderPath_GUID(FOLDERID_ConnectionsFolder) 'Virtual
'    sLog.AppendLine "FOLDERID_Contacts = " & GetKnownFolderPath_GUID(FOLDERID_Contacts)
'    'sLog.AppendLine "FOLDERID_ControlPanelFolder = " & GetKnownFolderPath_GUID(FOLDERID_ControlPanelFolder) 'Virtual
'    sLog.AppendLine "FOLDERID_Cookies = " & GetKnownFolderPath_GUID(FOLDERID_Cookies)
'    sLog.AppendLine "FOLDERID_Desktop = " & GetKnownFolderPath_GUID(FOLDERID_Desktop)
'    sLog.AppendLine "FOLDERID_DeviceMetadataStore = " & GetKnownFolderPath("{5CE4A5E9-E4EB-479D-B89F-130C02886155}")
'    sLog.AppendLine "FOLDERID_Documents = " & GetKnownFolderPath_GUID(FOLDERID_Documents)
'    sLog.AppendLine "FOLDERID_DocumentsLibrary = " & GetKnownFolderPath("{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}")
'    sLog.AppendLine "FOLDERID_Downloads = " & GetKnownFolderPath_GUID(FOLDERID_Downloads)
'    sLog.AppendLine "FOLDERID_Favorites = " & GetKnownFolderPath_GUID(FOLDERID_Favorites)
'    sLog.AppendLine "FOLDERID_Fonts = " & GetKnownFolderPath_GUID(FOLDERID_Fonts)
'    'sLog.AppendLine "FOLDERID_Games = " & GetKnownFolderPath_GUID(FOLDERID_Games) 'Virtual, less than 10, 1803
'    sLog.AppendLine "FOLDERID_GameTasks = " & GetKnownFolderPath_GUID(FOLDERID_GameTasks)
'    sLog.AppendLine "FOLDERID_History = " & GetKnownFolderPath_GUID(FOLDERID_History)
'    'sLog.AppendLine "FOLDERID_HomeGroup = " & GetKnownFolderPath_GUID(FOLDERID_HomeGroup) 'Virtual, 7+
'    'sLog.AppendLine "FOLDERID_HomeGroupCurrentUser = " & GetKnownFolderPath("{9B74B6A3-0DFD-4f11-9E78-5F7800F2E772}") 'Virtual, 8+
'    sLog.AppendLine "FOLDERID_ImplicitAppShortcuts = " & GetKnownFolderPath("{BCB5256F-79F6-4CEE-B725-DC34E402FD46}")
'    sLog.AppendLine "FOLDERID_InternetCache = " & GetKnownFolderPath_GUID(FOLDERID_InternetCache)
'    'sLog.AppendLine "FOLDERID_InternetFolder = " & GetKnownFolderPath_GUID(FOLDERID_InternetFolder) 'Virtual
'    sLog.AppendLine "FOLDERID_Libraries = " & GetKnownFolderPath("{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}")
'    sLog.AppendLine "FOLDERID_Links = " & GetKnownFolderPath_GUID(FOLDERID_Links)
'    sLog.AppendLine "FOLDERID_LocalAppData = " & GetKnownFolderPath_GUID(FOLDERID_LocalAppData)
'    sLog.AppendLine "FOLDERID_LocalAppDataLow = " & GetKnownFolderPath_GUID(FOLDERID_LocalAppDataLow)
'    sLog.AppendLine "FOLDERID_LocalizedResourcesDir = " & GetKnownFolderPath_GUID(FOLDERID_LocalizedResourcesDir)
'    sLog.AppendLine "FOLDERID_Music = " & GetKnownFolderPath_GUID(FOLDERID_Music)
'    sLog.AppendLine "FOLDERID_MusicLibrary = " & GetKnownFolderPath("{2112AB0A-C86A-4FFE-A368-0DE96E47012E}")
'    sLog.AppendLine "FOLDERID_NetHood = " & GetKnownFolderPath_GUID(FOLDERID_NetHood)
'    'sLog.AppendLine "FOLDERID_NetworkFolder = " & GetKnownFolderPath_GUID(FOLDERID_NetworkFolder) 'Virtual
'    If OSver.MajorMinor >= 10 And OSver.ReleaseId >= 1703 Then sLog.AppendLine "FOLDERID_Objects3D = " & GetKnownFolderPath("{31C0DD25-9439-4F12-BF41-7FF4EDA38722}") '10, 1703+
'    sLog.AppendLine "FOLDERID_OriginalImages = " & GetKnownFolderPath_GUID(FOLDERID_OriginalImages)
'    sLog.AppendLine "FOLDERID_PhotoAlbums = " & GetKnownFolderPath_GUID(FOLDERID_PhotoAlbums)
'    sLog.AppendLine "FOLDERID_Pictures = " & GetKnownFolderPath_GUID(FOLDERID_Pictures)
'    sLog.AppendLine "FOLDERID_PicturesLibrary = " & GetKnownFolderPath("{A990AE9F-A03B-4E80-94BC-9912D7504104}")
'    sLog.AppendLine "FOLDERID_Playlists = " & GetKnownFolderPath_GUID(FOLDERID_Playlists)
'    'sLog.AppendLine "FOLDERID_PrintersFolder = " & GetKnownFolderPath_GUID(FOLDERID_PrintersFolder) 'Virtual
'    sLog.AppendLine "FOLDERID_PrintHood = " & GetKnownFolderPath_GUID(FOLDERID_PrintHood)
'    sLog.AppendLine "FOLDERID_Profile = " & GetKnownFolderPath_GUID(FOLDERID_Profile)
'    sLog.AppendLine "FOLDERID_ProgramData = " & GetKnownFolderPath_GUID(FOLDERID_ProgramData)
'    sLog.AppendLine "FOLDERID_ProgramFiles = " & GetKnownFolderPath_GUID(FOLDERID_ProgramFiles)
'    sLog.AppendLine "FOLDERID_ProgramFilesCommon = " & GetKnownFolderPath_GUID(FOLDERID_ProgramFilesCommon)
'    'sLog.AppendLine "FOLDERID_ProgramFilesCommonX64 = " & GetKnownFolderPath_GUID(FOLDERID_ProgramFilesCommonX64) '(not supported for 32-bit applications)
'    sLog.AppendLine "FOLDERID_ProgramFilesCommonX86 = " & GetKnownFolderPath_GUID(FOLDERID_ProgramFilesCommonX86)
'    'sLog.AppendLine "FOLDERID_ProgramFilesX64 = " & GetKnownFolderPath_GUID(FOLDERID_ProgramFilesX64) '(not supported for 32-bit applications)
'    sLog.AppendLine "FOLDERID_ProgramFilesX86 = " & GetKnownFolderPath_GUID(FOLDERID_ProgramFilesX86)
'    sLog.AppendLine "FOLDERID_Programs = " & GetKnownFolderPath_GUID(FOLDERID_Programs)
'    sLog.AppendLine "FOLDERID_Public = " & GetKnownFolderPath_GUID(FOLDERID_Public)
'    sLog.AppendLine "FOLDERID_PublicDesktop = " & GetKnownFolderPath_GUID(FOLDERID_PublicDesktop)
'    sLog.AppendLine "FOLDERID_PublicDocuments = " & GetKnownFolderPath_GUID(FOLDERID_PublicDocuments)
'    sLog.AppendLine "FOLDERID_PublicDownloads = " & GetKnownFolderPath_GUID(FOLDERID_PublicDownloads)
'    sLog.AppendLine "FOLDERID_PublicGameTasks = " & GetKnownFolderPath_GUID(FOLDERID_PublicGameTasks)
'    sLog.AppendLine "FOLDERID_PublicLibraries = " & GetKnownFolderPath("{48DAF80B-E6CF-4F4E-B800-0E69D84EE384}")
'    sLog.AppendLine "FOLDERID_PublicMusic = " & GetKnownFolderPath_GUID(FOLDERID_PublicMusic)
'    sLog.AppendLine "FOLDERID_PublicPictures = " & GetKnownFolderPath_GUID(FOLDERID_PublicPictures)
'    sLog.AppendLine "FOLDERID_PublicRingtones = " & GetKnownFolderPath("{E555AB60-153B-4D17-9F04-A5FE99FC15EC}")
'    If OSver.MajorMinor >= 6.2 Then sLog.AppendLine "FOLDERID_PublicUserTiles = " & GetKnownFolderPath("{0482af6c-08f1-4c34-8c90-e17ec98b1e17}") '8+
'    sLog.AppendLine "FOLDERID_PublicVideos = " & GetKnownFolderPath_GUID(FOLDERID_PublicVideos)
'    sLog.AppendLine "FOLDERID_QuickLaunch = " & GetKnownFolderPath_GUID(FOLDERID_QuickLaunch)
'    sLog.AppendLine "FOLDERID_Recent = " & GetKnownFolderPath_GUID(FOLDERID_Recent)
'    If OSver.MajorMinor <= 6 Then sLog.AppendLine "FOLDERID_RecordedTV = " & GetKnownFolderPath_GUID(FOLDERID_RecordedTV) 'Vista-
'    sLog.AppendLine "FOLDERID_RecordedTVLibrary = " & GetKnownFolderPath("{1A6FDBA2-F42D-4358-A798-B74D745926C5}")
'    'sLog.AppendLine "FOLDERID_RecycleBinFolder = " & GetKnownFolderPath_GUID(FOLDERID_RecycleBinFolder) 'cannot be obtained in this way
'    sLog.AppendLine "FOLDERID_ResourceDir = " & GetKnownFolderPath_GUID(FOLDERID_ResourceDir)
'    sLog.AppendLine "FOLDERID_Ringtones = " & GetKnownFolderPath("{C870044B-F49E-4126-A9C3-B52A1FF411E8}")
'    If OSver.MajorMinor >= 6.2 Then sLog.AppendLine "FOLDERID_RoamedTileImages = " & GetKnownFolderPath("{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}") '8+
'    sLog.AppendLine "FOLDERID_RoamingAppData = " & GetKnownFolderPath_GUID(FOLDERID_RoamingAppData)
'    If OSver.MajorMinor >= 6.2 Then sLog.AppendLine "FOLDERID_RoamingTiles = " & GetKnownFolderPath("{00BCFC5A-ED94-4e48-96A1-3F6217F21990}") '8+
'    sLog.AppendLine "FOLDERID_SampleMusic = " & GetKnownFolderPath_GUID(FOLDERID_SampleMusic)
'    sLog.AppendLine "FOLDERID_SamplePictures = " & GetKnownFolderPath_GUID(FOLDERID_SamplePictures)
'
'    If OSver.MajorMinor <= 6.3 Then
'        sLog.AppendLine "FOLDERID_SamplePlaylists = " & GetKnownFolderPath_GUID(FOLDERID_SamplePlaylists) 'Windows 8.1- (check this !!!)
'    End If
'
'    sLog.AppendLine "FOLDERID_SampleVideos = " & GetKnownFolderPath_GUID(FOLDERID_SampleVideos)
'    sLog.AppendLine "FOLDERID_SavedGames = " & GetKnownFolderPath_GUID(FOLDERID_SavedGames)
'    'sLog.AppendLine "FOLDERID_SavedPictures = " & GetKnownFolderPath("{3B193882-D3AD-4eab-965A-69829D1FB59F}") 'cannot be obtained for some reason (???)
'    'sLog.AppendLine "FOLDERID_SavedPicturesLibrary = " & GetKnownFolderPath("{E25B5812-BE88-4bd9-94B0-29233477B6C3}") 'cannot be obtained for some reason (???)
'    sLog.AppendLine "FOLDERID_SavedSearches = " & GetKnownFolderPath_GUID(FOLDERID_SavedSearches)
'    If OSver.MajorMinor >= 6.2 Then sLog.AppendLine "FOLDERID_Screenshots = " & GetKnownFolderPath("{b7bede81-df94-4682-a7d8-57a52620b86f}") '8+
'    'sLog.AppendLine "FOLDERID_SEARCH_CSC = " & GetKnownFolderPath_GUID(FOLDERID_SEARCH_CSC) 'Virtual
'    'sLog.AppendLine "FOLDERID_SEARCH_MAPI = " & GetKnownFolderPath_GUID(FOLDERID_SEARCH_MAPI) 'Virtual
'    If OSver.MajorMinor >= 6.3 Then sLog.AppendLine "FOLDERID_SearchHistory = " & GetKnownFolderPath("{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}") '8.1+
'    'sLog.AppendLine "FOLDERID_SearchHome = " & GetKnownFolderPath_GUID(FOLDERID_SearchHome) 'Virtual
'    If OSver.MajorMinor >= 6.3 Then sLog.AppendLine "FOLDERID_SearchTemplates = " & GetKnownFolderPath("{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}") '8.1+
'    sLog.AppendLine "FOLDERID_SendTo = " & GetKnownFolderPath_GUID(FOLDERID_SendTo)
'
'    If OSver.MajorMinor <= 6.3 Then
'        If OSver.MajorMinor >= 6.1 Then sLog.AppendLine "FOLDERID_SidebarDefaultParts = " & GetKnownFolderPath_GUID(FOLDERID_SidebarDefaultParts) '7+, 10- (check this !!!)
'        If OSver.MajorMinor >= 6.1 Then sLog.AppendLine "FOLDERID_SidebarParts = " & GetKnownFolderPath_GUID(FOLDERID_SidebarParts) '7+, 10- (check this !!!)
'    End If
'
'    If OSver.MajorMinor >= 6.3 Then sLog.AppendLine "FOLDERID_SkyDrive = " & GetKnownFolderPath("{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}") '8.1+
'    If OSver.MajorMinor >= 6.3 Then sLog.AppendLine "FOLDERID_SkyDriveCameraRoll = " & GetKnownFolderPath("{767E6811-49CB-4273-87C2-20F355E1085B}") '8.1+
'    If OSver.MajorMinor >= 6.3 Then sLog.AppendLine "FOLDERID_SkyDriveDocuments = " & GetKnownFolderPath("{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}") '8.1+
'    If OSver.MajorMinor >= 6.3 Then sLog.AppendLine "FOLDERID_SkyDrivePictures = " & GetKnownFolderPath("{339719B5-8C47-4894-94C2-D8F77ADD44A6}") '8.1+
'    sLog.AppendLine "FOLDERID_StartMenu = " & GetKnownFolderPath_GUID(FOLDERID_StartMenu)
'    sLog.AppendLine "FOLDERID_Startup = " & GetKnownFolderPath_GUID(FOLDERID_Startup)
'    'sLog.AppendLine "FOLDERID_SyncManagerFolder = " & GetKnownFolderPath_GUID(FOLDERID_SyncManagerFolder) 'Virtual
'    'sLog.AppendLine "FOLDERID_SyncResultsFolder = " & GetKnownFolderPath_GUID(FOLDERID_SyncResultsFolder) 'Virtual
'    'sLog.AppendLine "FOLDERID_SyncSetupFolder = " & GetKnownFolderPath_GUID(FOLDERID_SyncSetupFolder) 'Virtual
'    sLog.AppendLine "FOLDERID_System = " & GetKnownFolderPath_GUID(FOLDERID_System)
'    sLog.AppendLine "FOLDERID_SystemX86 = " & GetKnownFolderPath_GUID(FOLDERID_SystemX86)
'    sLog.AppendLine "FOLDERID_Templates = " & GetKnownFolderPath_GUID(FOLDERID_Templates)
'    sLog.AppendLine "FOLDERID_UserPinned = " & GetKnownFolderPath("{9E3995AB-1F9C-4F13-B827-48B24B6C7174}")
'    sLog.AppendLine "FOLDERID_UserProfiles = " & GetKnownFolderPath_GUID(FOLDERID_UserProfiles)
'    sLog.AppendLine "FOLDERID_UserProgramFiles = " & GetKnownFolderPath("{5CD7AEE2-2219-4A67-B85D-6C9CE15660CB}")
'    sLog.AppendLine "FOLDERID_UserProgramFilesCommon = " & GetKnownFolderPath("{BCBD3057-CA5C-4622-B42D-BC56DB0AE516}")
'    'sLog.AppendLine "FOLDERID_UsersFiles = " & GetKnownFolderPath_GUID(FOLDERID_UsersFiles) 'Virtual
'    'sLog.AppendLine "FOLDERID_UsersLibraries = " & GetKnownFolderPath("{A302545D-DEFF-464b-ABE8-61C8648D939B}") 'Virtual
'    sLog.AppendLine "FOLDERID_Videos = " & GetKnownFolderPath_GUID(FOLDERID_Videos)
'    sLog.AppendLine "FOLDERID_VideosLibrary = " & GetKnownFolderPath("{491E922F-5643-4AF4-A7EB-4E7A138D8174}")
'    sLog.AppendLine "FOLDERID_Windows = " & GetKnownFolderPath_GUID(FOLDERID_Windows)
'
'    Exit Sub
'ErrorHandler:
'    ErrorMsg Err, "GetSpecialFolders_Vista_2"
'    If inIDE Then Stop: Resume Next
'End Sub

Private Sub GetSpecialFolders_Vista(sLog As clsStringBuilder)
    On Error GoTo ErrorHandler:
    Dim kfid As UUID
    Dim nCount As Long
    Dim pakfid As Long
    Dim i As Long
    Dim ptr As Long
    Dim Flags As Long
    Dim lpPath As Long
    Dim sPath$, sName$
    Dim aPath() As String
    
    Dim pKFM As KnownFolderManager
    Set pKFM = New KnownFolderManager
    
    Dim pKF As IKnownFolder
    Dim pKFD As KNOWNFOLDER_DEFINITION
    
    Dim pItem As IShellItem
    Dim penum1 As IEnumShellItems
    Dim pChild As IShellItem
    Dim pcl As Long
    
    Call pKFM.GetFolderIds(pakfid, nCount)
    
    If nCount > 0 Then
        
        ptr = pakfid
        ReDim aPath(nCount - 1)
        
        For i = 1 To nCount
            memcpy kfid, ByVal ptr, LenB(kfid)  'array[idx] -> UUID
            ptr = ptr + LenB(kfid)
            
            Call pKFM.GetFolder(kfid, pKF)  'UUID -> IKnownFolder
            
            If Not (pKF Is Nothing) Then
                
                sName = "": sPath = ""
                
                Call pKF.GetFolderDefinition(pKFD)   'IKnownFolder -> KNOWNFOLDER_DEFINITION
                
                If pKFD.pszName <> 0 Then
                    sName = BStrFromLPWStr(pKFD.pszName)        'get name
                End If
                
                'FreeKnownFolderDefinitionFields pKFD
                
                If sName = "RecycleBinFolder" Then
                    On Error Resume Next
                    pKF.GetShellItem KF_FLAG_DEFAULT, IID_IShellItem, pItem
                    On Error GoTo ErrorHandler:

                    If Not (pItem Is Nothing) Then

                        'special method: retrieve path by listing first child item

                        pItem.BindToHandler ByVal 0&, BHID_EnumItems, IID_IEnumShellItems, penum1

                        If Not (penum1 Is Nothing) Then
                            If penum1.Next(1&, pChild, pcl) = S_OK Then

                                pChild.GetAttributes SFGAO_FILESYSTEM, Flags
                                If Flags And SFGAO_FILESYSTEM Then
                                    pChild.GetDisplayName SIGDN_FILESYSPATH, lpPath
                                    sPath = BStrFromLPWStr(lpPath, True)
                                    sPath = GetParentDir(sPath)
                                End If
                            End If
                        End If
                    End If
                Else
                    Flags = (KF_FLAG_SIMPLE_IDLIST Or KF_FLAG_DONT_VERIFY Or KF_FLAG_DEFAULT_PATH Or KF_FLAG_NOT_PARENT_RELATIVE)
                    On Error Resume Next
                    Call pKF.GetPath(Flags, lpPath)      'IKnownFolder -> physical path
                    If lpPath <> 0 Then
                        sPath = BStrFromLPWStr(lpPath, True)
                    End If
                    On Error GoTo ErrorHandler:
                End If
                
                If Len(sPath) <> 0 Then
                    aPath(i - 1) = sName & " = " & sPath
                End If
            End If
        Next
        
        CoTaskMemFree pakfid
    End If
    
    Set pKFM = Nothing
    
    If AryPtr(aPath) Then
        CompressArray aPath
        QuickSort aPath, 0, UBound(aPath)
        For i = 0 To UBound(aPath)
            sLog.AppendLine aPath(i)
        Next
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetSpecialFolders_Vista"
    If inIDE Then Stop: Resume Next
End Sub

Public Function CreateLogFile() As String
    Dim sLog As clsStringBuilder
    Dim i&, j&, sProcessList$
    Dim lNumProcesses&
    Dim sProcessName$
    Dim Col As New Collection, cnt&
    Dim sTmp$
    Dim bStadyMakeLog As Boolean
    Dim aPos() As Variant, aNames() As String
    Dim sScanMode As String
    Dim dModule As clsTrickHashTable
    Dim aModules() As String
    Dim sModule As String
    Dim sModuleList As String
    Dim bMicrosoft As Boolean
    
    On Error GoTo ErrorHandler:
    
    AppendErrorLogCustom "frmMain.CreateLogFile - Begin"
    
    Set sLog = New clsStringBuilder 'speed optimization for huge logs
    
    If Not bLogProcesses Then GoTo MakeLog
    
    If Not bAutoLogSilent Then
        lNumProcesses = GetProcesses(gProcess)
    Else
        If AryPtr(gProcess) Then
            lNumProcesses = UBound(gProcess) + 1
        End If
    End If
    
    If (lNumProcesses > 0) Then
    
        For i = 0 To UBound(gProcess)
            
            sProcessName = gProcess(i).Path
            
            If Len(gProcess(i).Path) = 0 Then
                If bIgnoreAllWhitelists Or Not IsDefaultSystemProcess(gProcess(i).pid, gProcess(i).Name, gProcess(i).Path) Then
                    sProcessName = gProcess(i).Name
                End If
            End If
            
            If Len(sProcessName) <> 0 Then
                If Not isCollectionKeyExists(sProcessName, Col) Then
                    Col.Add 1&, sProcessName          ' item - count, key - name of process
                Else
                    cnt = Col.Item(sProcessName)      ' replacing item of collection
                    Col.Remove (sProcessName)
                    Col.Add cnt + 1&, sProcessName    ' increase count of identical processes
                End If
            End If
        Next
    End If
    
    'sProcessList = "Running processes:" & vbCrLf
    sProcessList = Translate(29) & ":" & vbCrLf
    
    'sProcessList = sProcessList & "Number | Path" & vbCrLf
    If bAdditional Then
        sProcessList = sProcessList & "  PID | " & Translate(1021) & vbCrLf
    Else
        sProcessList = sProcessList & Translate(1020) & " | " & Translate(1021) & vbCrLf
    End If
    
    If bLogModules Then
        'Additional mode => PID | Process Name
        
        ReDim aPos(UBound(gProcess)), aNames(UBound(gProcess))
        
        For i = 0 To UBound(gProcess)
            aPos(i) = i
            If Len(gProcess(i).Path) = 0 Then
                gProcess(i).Path = gProcess(i).Name
            End If
            aNames(i) = gProcess(i).Path
        Next
        
        QuickSortSpecial aNames, aPos, 0, UBound(gProcess)
        
        For i = 0 To UBound(aPos)
            With gProcess(aPos(i))
                '// TODO: add 'is microsoft' check and mark
                sProcessList = sProcessList & Right$("     " & .pid & "  ", 8) & .Path & vbCrLf
            End With
        Next
        
    Else
        'Normal mode => Number of processes | Process Name
    
        ' Sort using positions array method (Key - Process Path).
        ReDim ProcLog(Col.Count) As MY_PROC_LOG
        ReDim aPos(Col.Count), aNames(Col.Count)
        
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
                sProcessList = sProcessList & Right$("   " & .Number & "  ", 6) & IIf(.IsMicrosoft, "(Microsoft) ", "") & .ProcName & vbCrLf
            End With
        Next
    End If
    
    sProcessList = sProcessList & vbCrLf
    
    'show all PIDs in debug. mode
    If bDebug Or bDebugToFile Then
        If lNumProcesses Then
            sTmp = ""
            For i = 0 To UBound(gProcess)
                sTmp = sTmp & gProcess(i).pid & " | " & IIf(Len(gProcess(i).Path) <> 0, gProcess(i).Path, gProcess(i).Name) & vbCrLf
            Next
            AppendErrorLogCustom sTmp
            sTmp = ""
        End If
    End If
    
    If bLogModules Then
        'Loaded modules
        
        Set dModule = New clsTrickHashTable
        
        For i = 0 To UBound(gProcess)
        
            If gProcess(i).pid <> 0 And gProcess(i).pid <> 4 Then
        
                aModules = EnumModules64(gProcess(i).pid)
        
                If AryItems(aModules) Then
                    For j = 0 To UBound(aModules)
                        If Not dModule.Exists(aModules(j)) Then
                            dModule.Add aModules(j), "[" & CStr(gProcess(i).pid) & "]"
                        Else
                            dModule(aModules(j)) = dModule(aModules(j)) & " [" & CStr(gProcess(i).pid) & "]"
                        End If
                    Next
                End If
            End If
        Next
        
        If dModule.Count > 0 Then
            ReDim aPos(dModule.Count - 1)
            ReDim aNames(dModule.Count - 1)
        
            For i = 0 To dModule.Count - 1
                aPos(i) = i
                aNames(i) = dModule.Keys(i)
            Next
            
            QuickSortSpecial aNames, aPos, 0, UBound(aNames)
            
            '// TODO: add command line extraction via PEB
            
            'Loaded modules: Path | PID | Command line
            sModuleList = Translate(21) & ":" & vbCrLf & Translate(1021) & " | PID" & vbCrLf 'translate(1070)
            
            For i = 0 To UBound(aPos)
                If Not bAutoLogSilent Then
                    DoEvents
                    UpdateProgressBar "ModuleList", CStr(i + 1) & " / " & CStr(UBound(aPos) + 1)
                End If
                sModule = dModule.Keys(aPos(i))
                bMicrosoft = IsMicrosoftFile(sModule)
                
                If (Not bMicrosoft) Or Not bHideMicrosoft Then
                    sModuleList = sModuleList & sModule & vbTab & vbTab & vbTab & dModule.Items(aPos(i)) & vbCrLf
                End If
            Next
            
        End If
            
        Set dModule = Nothing
    End If
    
    '------------------------------
MakeLog:
    bStadyMakeLog = True
    
    UpdateProgressBar "Report"
    
    'UpdateProgressBar "Finish"
    'DoEvents
    
    sLog.Append ChrW$(-257) & "Logfile of " & AppVerPlusName & vbCrLf & vbCrLf ' + BOM UTF-16 LE
    
    sLog.Append MakeLogHeader()
    
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
    
    If (Not bLogProcesses) Or bLogModules Or bAdditional Or bLogEnvVars Or (Not bHideMicrosoft) Or bIgnoreAllWhitelists Then
        If bAdditional Then
            sScanMode = "Additional"
        End If
        If Not bLogProcesses Then
            sScanMode = sScanMode & "; Do not log Processes"
        End If
        If bLogModules Then
            sScanMode = sScanMode & "; Log Loaded Modules"
        End If
        If bLogEnvVars Then
            sScanMode = sScanMode & "; Environment variables"
        End If
        If Not bHideMicrosoft Then
            sScanMode = sScanMode & "; Do not Hide Microsoft"
        End If
        If bIgnoreAllWhitelists Then
            sScanMode = sScanMode & "; Ignore ALL Whitelists"
        End If
        If Left(sScanMode, 2) = "; " Then sScanMode = Mid$(sScanMode, 3)
        
        sLog.Append "Scan mode: " & sScanMode & vbCrLf
    End If
    
'    If OSver.IsSystemCaseSensitive Then
'        sLog.Append "Filenames: is in case sensitive mode" & vbCrLf
'    End If
    
    If bLogEnvVars Then
        
        Dim pEnv As Long
        Dim pEnvNext As Long
        Dim strEnv As String
        Dim sEnvName As String
        Dim sEnvValue As String
        Dim bAddEV As Boolean
        Dim varCnt As Long
        Dim aValues() As String
        Dim pos As Long
        
        'Dim varDict As clsTrickHashTable
        'Set varDict = New clsTrickHashTable
        
        sLog.Append vbCrLf & "Environment variables:" & vbCrLf & vbCrLf
        
        'get System EV
        sLog.AppendLine "[System]"
        For i = 1 To Reg.EnumValuesToArray(HKLM, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", aValues, False)
            sEnvName = aValues(i)
            sEnvValue = Reg.GetString(HKLM, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", sEnvName, False)
            sLog.AppendLine sEnvName & " = " & sEnvValue
            'If Not varDict.Exists(sEnvName) Then varDict.Add sEnvName, sEnvValue
        Next
        
        'get User EV
        sLog.AppendLine
        sLog.AppendLine "[User]"
        varCnt = Reg.EnumValuesToArray(HKCU, "Environment", aValues, False)
        For i = 1 To varCnt
            sEnvName = aValues(i)
            sEnvValue = Reg.GetString(HKCU, "Environment", sEnvName, False)
            sLog.AppendLine sEnvName & " = " & sEnvValue
            'If Not varDict.Exists(sEnvName) Then varDict.Add sEnvName, sEnvValue
        Next
        If varCnt = 0 Then
            sLog.AppendLine "No variables."
        End If
        
        'get process EV
        sLog.AppendLine
        sLog.AppendLine "[Current process]"
        varCnt = 0
        
        pEnv = GetEnvironmentStrings()
        pEnvNext = pEnv
        
        If pEnv <> 0 Then
            Do
                strEnv = StringFromPtrW(pEnvNext)
                
                If Len(strEnv) <> 0 Then
                    pos = InStr(2, strEnv, "=")
                    If pos = 0 Then pos = InStr(1, strEnv, "=")
                    If pos <> 0 Then
                        sEnvName = Left$(strEnv, pos - 1)
                        sEnvValue = Mid$(strEnv, pos + 1)
                        'bAddEV = True
                        'If varDict.Exists(sEnvName) Then
                        '    If varDict(sEnvName) = sEnvValue Then bAddEV = False
                        'End If
                        'If bAddEV Then
                            sLog.AppendLine sEnvName & " = " & sEnvValue
                            varCnt = varCnt + 1
                        'End If
                    End If
                    pEnvNext = pEnvNext + LenB(strEnv) + 2
                End If
            Loop Until Len(strEnv) = 0
            
            FreeEnvironmentStrings pEnv
        End If
        If varCnt = 0 Then
            'sLog.AppendLine "All variables are identical to System/User."
        End If
        
        'Set varDict = Nothing
        
        sLog.AppendLine
        sLog.AppendLine "Special folders:"
        sLog.AppendLine

        If OSver.IsWindowsVistaOrGreater Then
            GetSpecialFolders_Vista sLog
        Else
            'GetSpecialFolders_XP sLog
            GetSpecialFolders_XP_2 sLog
        End If
        
        '// TODO
        '
        'Append with:
        'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders
        'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders
        'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders
        'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders
        '
        'because Microsoft doesn't care about its own "Best practice", mentioned in MSDN,
        'so Special Folders retrieved by CLSID become completely useless, e.g. for "Desktop" redirection in Win 10.
        '
    End If
    
    sLog.Append vbCrLf & sProcessList
    
    If bLogModules Then
        sLog.Append sModuleList & vbCrLf & vbCrLf
    End If
    
    ' -----> MAIN Sections <------
    
    ' in /silentautolog mode result screen is not fill due to speed optimization
    If (Not bAutoLogSilent) And frmMain.lstResults.ListCount = 0 Then
        sLog.AppendLine Translate(1004) 'No suspicious items found!
    Else
        If AryItems(HitSorted) Then
            For i = 0 To UBound(HitSorted)
                ' Adding empty lines beetween sections (cancelled)
                'sPrefix = rtrim$(Splitsafe(HitSorted(i), "-")(0))
                'If sPrefixLast <> "" And sPrefixLast <> sPrefix Then sLog = sLog & vbCrLf
                'sPrefixLast = sPrefix
                sLog.Append HitSorted(i) & vbCrLf
            Next
        End If
    End If
    
    ' ----------------------------
    
    Dim IgnoreCnt&
    IgnoreCnt = RegReadHJT("IgnoreNum", "0")
    If IgnoreCnt <> 0 Then
        If bSkipIgnoreList Or bLoadDefaults Then
            ' "Warning: Ignore list contains " & IgnoreCnt & " items, but they are displayed in log, because /default (/skipIgnoreList) switch is used." & vbCrLf
            sLog.Append vbCrLf & vbCrLf & Replace$(Replace$(Translate(1017), "[]", IgnoreCnt), "[*]", IIf(bLoadDefaults, "default", "skipIgnoreList")) & vbCrLf
        Else
            ' "Warning: Ignore list contains " & IgnoreCnt & " items." & vbCrLf
            sLog.Append vbCrLf & vbCrLf & Replace$(Translate(1011), "[]", IgnoreCnt) & vbCrLf
        End If
    End If
    If Not bScanExecuted Then
        ' "Warning: General scanning was not performed." & vbCrLf
        sLog.Append vbCrLf & vbCrLf & Translate(1012) & vbCrLf
    End If
    Dim SignMsg As String
    If Not isEDS_Work(SignMsg) Then
        sLog.Append vbCrLf & vbCrLf & BuildPath(sWinDir, "system32\ntdll.dll") & " file doesn't pass digital signature verification. Error: " & SignMsg & vbCrLf
    End If
    If Not isEDS_CatExist() Then
        sLog.Append "Windows security catalogue is empty!" & vbCrLf
    End If
    
    If Len(EndReport) <> 0 Then
        sLog.AppendLine vbCrLf & EndReport
    End If
    EndReport = vbNullString
    
    'Append by Error Log
    If 0 <> Len(ErrReport) Then
        sLog.Append vbCrLf & vbCrLf & "Debug information:" & vbCrLf & ErrReport & vbCrLf
        '& vbCrLf & "CmdLine: " & AppPath(True) & " " & g_sCommandLine
    End If
    
    Dim b()     As Byte
    
    If bDebugToFile Then
        If g_hDebugLog <> 0 Then
            b() = vbCrLf & vbCrLf & "Contents of the main logfile:" & vbCrLf & vbCrLf & sLog.ToString & vbCrLf
            PutW g_hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
        End If
    End If
    
    If 0 <> ErrLogCustomText.Length Then
        sLog.Append vbCrLf & vbCrLf & "Trace information:" & vbCrLf & ErrLogCustomText.ToString & vbCrLf
    End If
    
    If bAutoLog Then Perf.EndTime = GetTickCount()
    sLog.Append vbCrLf & "--" & vbCrLf & "End of file - " & "Time spent: " & ((Perf.EndTime - Perf.StartTime) \ 100) / 10 & " sec. - "
    
    If bDebugToFile Then
        If g_hDebugLog <> 0 Then
            b() = vbCrLf & "--" & vbCrLf & "End of file - " & "Time spent: " & ((Perf.EndTime - Perf.StartTime) \ 100) / 10 & " sec."
            PutW g_hDebugLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
            CloseW g_hDebugLog, True: g_hDebugLog = 0
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
    
    Set sLog = Nothing
    
    AppendErrorLogCustom "frmMain.CreateLogFile - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "frmMain_CreateLogFile"
    If inIDE Then Stop: Resume Next
    If Not bStadyMakeLog Then GoTo MakeLog
    Set sLog = Nothing
End Function

Private Function isEDS_CatExist() As Boolean
    'check is it have at least 10 cat. files (141 is in Win2k, 30 - is in XP SP2)
    'c:\Windows\System32\catroot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}
    
    Const NUM_REQUIRED = 10&
    
    Dim hFind As Long
    Dim cnt As Long
    Dim fd As WIN32_FIND_DATA
        
    hFind = FindFirstFile(StrPtr(BuildPath(sWinSysDir, "catroot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\*.cat")), fd)
    If hFind <> 0& Then
        Do Until FindNextFile(hFind, fd) = 0& Or cnt >= NUM_REQUIRED
            cnt = cnt + 1
        Loop
        FindClose hFind
    End If
    
    isEDS_CatExist = (cnt >= NUM_REQUIRED)
End Function

Public Function MakeLogHeader() As String

    Dim TimeCreated As String
    Dim bSPOld As Boolean
    Dim sUTC As String
    Dim sText As String
    
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
    
    sText = "Platform:  " & OSver.Bitness & " " & OSver.OSName & IIf(OSver.Edition <> "", " (" & OSver.Edition & ")", "") & ", " & _
            OSver.Major & "." & OSver.Minor & "." & OSver.Build & "." & OSver.Revision & _
            IIf(OSver.ReleaseId <> 0, " (ReleaseId: " & OSver.ReleaseId & ")", "") & ", " & _
            "Service Pack: " & OSver.SPVer & IIf(bSPOld, " <=== Attention! (outdated SP)", "") & _
            IIf(OSver.MajorMinor <> OSver.MajorMinorNTDLL And OSver.MajorMinorNTDLL <> 0, " (ntdll.dll = " & OSver.NtDllVersion & ")", "") & _
            vbCrLf
    
    '," & vbTab & "Uptime: " & TrimSeconds(GetSystemUpTime()) & " h/m" & vbCrLf
            
    sText = sText & "Time:      " & TimeCreated & " (" & sUTC & ")" & vbCrLf
    sText = sText & "Language:  " & "OS: " & OSver.LangSystemNameFull & " (" & "0x" & Hex$(OSver.LangSystemCode) & "). " & _
            "Display: " & OSver.LangDisplayNameFull & " (" & "0x" & Hex$(OSver.LangDisplayCode) & "). " & _
            "Non-Unicode: " & OSver.LangNonUnicodeNameFull & " (" & "0x" & Hex$(OSver.LangNonUnicodeCode) & ")" & vbCrLf
    
    If OSver.MajorMinor >= 6 Then
        sText = sText & "Elevated:  " & IIf(OSver.IsElevated, "Yes", "No") & vbCrLf  '& vbTab & "IL: " & OSver.GetIntegrityLevel & vbCrLf
    End If
    
    sText = sText & "Ran by:    " & OSver.UserName & vbTab & "(group: " & OSver.UserType & ") on " & OSver.ComputerName & _
        ", " & IIf(bDebugMode, "(SID: " & OSver.SID_CurrentProcess & ") ", "") & "FirstRun: " & IIf(bFirstRebootScan, "yes", "no") & _
        IIf(OSver.IsLocalSystemContext, " <=== Attention! ('Local System' account)", "") & vbCrLf & vbCrLf
        
    MakeLogHeader = sText
End Function

' Сортировка по Хоару. На вход - массив j(), на выходе массив k() с индексами массива j в отсортированном порядке + отсортированный массив.
' Алгоритм особо отлично подходит для сортировки User type arrays по любому из полей.
Public Sub QuickSortSpecial(j() As String, k() As Variant, ByVal low As Long, ByVal high As Long)
    On Error GoTo ErrorHandler:
    Dim i As Long, L As Long, M As String, wsp As String
    i = low: L = high: M = j((i + L) \ 2)
    Do Until i > L: Do While j(i) < M: i = i + 1: Loop: Do While j(L) > M: L = L - 1: Loop
        If (i <= L) Then wsp = j(i): j(i) = j(L): j(L) = wsp: wsp = k(i): k(i) = k(L): k(L) = wsp: i = i + 1: L = L - 1
    Loop
    If low < L Then QuickSortSpecial j, k, low, L
    If i < high Then QuickSortSpecial j, k, i, high
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
    Dim sSelItem As String
    'HitSorted() -> is a global array
    
    'save selected position on listbox
    If frmMain.lstResults.ListIndex <> -1 Then
        sSelItem = frmMain.lstResults.List(frmMain.lstResults.ListIndex)
    End If
    
    Dim iOldTop&
    iOldTop = frmMain.lstResults.TopIndex
    
    Erase HitSorted
    
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
    
    'restore selected position in listbox
    If Len(sSelItem) <> 0 Then
        For i = 0 To frmMain.lstResults.ListCount - 1
            If StrComp(frmMain.lstResults.List(i), sSelItem) = 0 Then
                frmMain.lstResults.ListIndex = i
                Exit For
            End If
        Next
    End If

    If iOldTop <> -1 Then frmMain.lstResults.TopIndex = iOldTop
    
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
    Dim SectName As String
    Dim nHit As Long
    Dim nSect As Long
    Dim nItemsSect As Long
    Dim nItemsHit As Long
    Dim pos As Long
    Dim i As Long
    Dim bComply As Boolean
    Dim cSectNames As Collection
    Set cSectNames = New Collection

    ReDim aDstArray(UBound(aSrcArray))
    
    cSectNames.Add "R0"
    cSectNames.Add "R1"
    cSectNames.Add "R2"
    cSectNames.Add "R3"
    cSectNames.Add "R4"
    cSectNames.Add "F0"
    cSectNames.Add "F1"
    cSectNames.Add "F2"
    cSectNames.Add "F3"
    For i = 1 To 22
        cSectNames.Add "O" & i
    Next
    cSectNames.Add "O23 - Service"
    cSectNames.Add "O23 - Driver"
    cSectNames.Add "O23"
    For i = 24 To LAST_CHECK_OTHER_SECTION_NUMBER '26
        cSectNames.Add "O" & i
    Next
    
    'Алгоритм сортировки:
    'Перечисляем префикс каждой из секций и ищем его в строках лога
    'Как только массив секции собран, сортируем его и результат сбрасываем в результирующий массив
    
    nItemsHit = 0
    
    For nSect = 1 To cSectNames.Count
        nItemsSect = 0
        For nHit = 0 To UBound(aSrcArray)
            If 0 <> Len(aSrcArray(nHit)) Then
                pos = InStr(aSrcArray(nHit), "-")
                If pos = 0 Then
                    If Not bAutoLog Then
                        MsgBoxW "Warning! Wrong format of hit line. Must include dash after the name of the section!" & vbCrLf & "Line: " & aSrcArray(nHit)
                    End If
                Else
                    bComply = False
                    
                    'если секция собирается по префиксу и частичному имени
                    If InStr(cSectNames(nSect), "-") <> 0 Then
                        If StrBeginWith(aSrcArray(nHit), cSectNames(nSect)) Then bComply = True
                    Else
                        'если секция собирается только по префиксу
                        SectName = Trim$(Left$(aSrcArray(nHit), pos - 1))
                        If SectName = cSectNames(nSect) Then bComply = True
                    End If
                
                    'строка лога соответствует критерию -> добавляем к секции
                    If bComply Then
                        ' Создаю временный массив этой секции для сортировки
                        nItemsSect = nItemsSect + 1
                        ReDim Preserve SectSorted(nItemsSect)
                        'переношу в SectSorted данные из aSrcArray
                        SectSorted(nItemsSect) = aSrcArray(nHit)
                        'а в исходном массиве убираю строку
                        aSrcArray(nHit) = vbNullString
                    End If
                End If
            End If
        Next
        ' Сборка секции завершена.
        If 0 <> nItemsSect Then
            ' Начало сортировки секции
            ' O1 не сортируем (hosts)
            If cSectNames(nSect) <> "O1" Then
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
    ReDim SectSorted(0)
    nItemsSect = 0
    For nHit = 0 To UBound(aSrcArray)
        If 0 <> Len(aSrcArray(nHit)) Then
            'собираем остатки в массив SectSorted
            nItemsSect = nItemsSect + 1
            ReDim Preserve SectSorted(nItemsSect)
            SectSorted(nItemsSect) = aSrcArray(nHit)
        End If
    Next
    If nItemsSect > 0 Then
        'сортируем его
        QuickSort SectSorted, 0, UBound(SectSorted)
        
        'и сбрасываем в конец результирующего массива
        For i = 0 To UBound(SectSorted)
            If Len(SectSorted(i)) <> 0 Then
                aDstArray(nItemsHit) = SectSorted(i)
                nItemsHit = nItemsHit + 1
            End If
        Next
    End If
    
    Set cSectNames = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "frmMain_SortSectionsOfResultListEx"
    If inIDE Then Stop: Resume Next
End Sub

' Add key to jump list without actually doing cure on it
' Returns 'true' if key/value is exist and was add to jump list
Public Function AddJumpRegistry( _
    JumpArray() As JUMP_ENTRY, _
    ActionType As ENUM_REG_ACTION_BASED, _
    ByVal lHive As ENUM_REG_HIVE, _
    ByVal sKey As String, _
    Optional sParam As String = "", _
    Optional vDefaultData As Variant = "", _
    Optional eRedirected As ENUM_REG_REDIRECTION = REG_NOTREDIRECTED, _
    Optional ParamType As ENUM_REG_VALUE_TYPE_RESTORE = REG_RESTORE_SAME) As Boolean
    
    'speed hack
    If bAutoLogSilent Then Exit Function
    
    Dim KeyFix() As FIX_REG_KEY
    AddRegToFix KeyFix, ActionType, lHive, sKey, sParam, vDefaultData, eRedirected, ParamType
    
    If AryPtr(KeyFix) Then
        AddJumpRegistry = True
    
        If AryPtr(JumpArray) Then
            ReDim Preserve JumpArray(UBound(JumpArray) + 1)
        Else
            ReDim JumpArray(0)
        End If
        With JumpArray(UBound(JumpArray))
            .Type = JUMP_ENTRY_REGISTRY
            .Registry = KeyFix
        End With
    End If
End Function

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
    
    'speed hack
    If bAutoLogSilent Then Exit Sub
    
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
                    
                    If (ActionType And BACKUP_KEY) Or (ActionType And REMOVE_KEY) Or (ActionType And REMOVE_KEY_IF_NO_VALUES) Or (ActionType And JUMP_KEY) Then
                        If Not Reg.KeyExists(lActualHive, sKey, Wow6432Redir) Then bNoItem = True
                        
                    ElseIf (ActionType And BACKUP_VALUE) Or (ActionType And REMOVE_VALUE) _
                      Or (ActionType And REMOVE_VALUE_IF_EMPTY) Or (ActionType And JUMP_VALUE) Then
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
    sSection As Variant, _
    Optional sParam As Variant = "", _
    Optional sDefaultData As Variant = "")
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent Then Exit Sub
    
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

' Add file to jump list without actually doing cure on it
' Returns 'true' if file exists and was add to jump list
Public Function AddJumpFile( _
    JumpArray() As JUMP_ENTRY, _
    ActionType As ENUM_FILE_ACTION_BASED, _
    sFilePath As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent Then Exit Function
    
    If Len(sFilePath) = 0 Then Exit Function
    
    Dim FileFix() As FIX_FILE
    AddFileToFix FileFix, ActionType, sFilePath
    
    If AryPtr(FileFix) Then
        AddJumpFile = True
    
        If AryPtr(JumpArray) Then
            ReDim Preserve JumpArray(UBound(JumpArray) + 1)
        Else
            ReDim JumpArray(0)
        End If
        With JumpArray(UBound(JumpArray))
            .Type = JUMP_ENTRY_FILE
            .File = FileFix
        End With
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "AddJumpFile", sFilePath
    If inIDE Then Stop: Resume Next
End Function

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
    
    'speed hack
    If bAutoLogSilent Then Exit Sub
    
    If Len(sFilePath) = 0 Then Exit Sub
    If FileMissing(sFilePath) Then Exit Sub
    
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
    
        'if no file to jump
        If ActionType And JUMP_FILE Then
            If Not FileExists(sFilePath) Then Exit Sub
        End If
    
        'if no folder to remove
        If ActionType And REMOVE_FOLDER Then
            If Not FolderExists(sFilePath) Then Exit Sub
        End If
        
        'if no folder to jump
        If ActionType And JUMP_FOLDER Then
            If Not FolderExists(sFilePath) Then Exit Sub
        End If
        
        'if folder is already exist
        If ActionType And CREATE_FOLDER Then
            If FolderExists(sFilePath) Then Exit Sub
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
    
    'speed hack
    If bAutoLogSilent Then Exit Sub
    
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

' Append results array with new custom record
Public Sub AddCustomToFix( _
    CustomArray() As FIX_CUSTOM, _
    ActionType As ENUM_CUSTOM_ACTION_BASED, _
    Optional sObjectName As String)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent Then Exit Sub
    
    If AryPtr(CustomArray) Then
        ReDim Preserve CustomArray(UBound(CustomArray) + 1)
    Else
        ReDim CustomArray(0)
    End If
    
    With CustomArray(UBound(CustomArray))
        .ActionType = ActionType
        .ObjectName = sObjectName
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddCustomToFix", ActionType
    If inIDE Then Stop: Resume Next
End Sub

' Append results array with new process record
Public Sub AddServiceToFix( _
    ServiceArray() As FIX_SERVICE, _
    ActionType As ENUM_SERVICE_ACTION_BASED, _
    sServiceName As String, _
    Optional sServiceDisplay As String = "", _
    Optional sImagePath As String = "", _
    Optional sDllPath As String = "", _
    Optional RunState As SERVICE_STATE)
    
    On Error GoTo ErrorHandler
    
    'speed hack
    If bAutoLogSilent Then Exit Sub
    
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
        .RunState = RunState
    End With
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "AddServiceToFix", ActionType, sServiceName, sServiceDisplay, sImagePath, sDllPath
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixIt(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    If result.CureType And SERVICE_BASED Then FixServiceHandler result
    If result.CureType And PROCESS_BASED Then FixProcessHandler result
    If result.CureType And FILE_BASED Then FixFileHandler result
    If result.CureType And (REGISTRY_BASED Or INI_BASED) Then FixRegistryHandler result
    If result.CureType And CUSTOM_BASED Then FixCustomHandler result
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixIt", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub


Public Sub FixCustomHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    If result.CureType And CUSTOM_BASED Then
        If AryPtr(result.Custom) Then
            For i = 0 To UBound(result.Custom)
                With result.Custom(i)
                    Select Case .ActionType
                    
                    Case CUSTOM_ACTION_O25
                        RemoveSubscriptionWMI result
                    
                    End Select
                End With
            Next
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixCustomHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub
                

Public Sub FixProcessHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim bParentProtected As Boolean
    
    If result.CureType And PROCESS_BASED Then
        If AryPtr(result.Process) Then
            For i = 0 To UBound(result.Process)
                With result.Process(i)
                    
                    'my parent and not explorer ?
                    bParentProtected = StrComp(.Path, MyParentProc.Path, 1) = 0 And Not StrEndWith(.Path, "explorer.exe")
                    
                    If Not IsSystemCriticalProcessPath(.Path) And Not bParentProtected Then
                        
                        If (.ActionType And USE_FEATURE_DISABLE) And g_bDelmodeDisabling Then
                            Exit Sub
                        End If
                    
                        If .ActionType And FREEZE_PROCESS Then
                        
                            PauseProcessByFile .Path
                        End If
                    
                        If .ActionType And KILL_PROCESS Then
                            
                            KillProcessByFile .Path, bForceMicrosoft:=True
                        End If
                        
                        If .ActionType And FREEZE_OR_KILL_PROCESS Then
                        
                            If Not PauseProcessByFile(.Path) Then
                                KillProcessByFile .Path, bForceMicrosoft:=True
                            End If
                        End If
                    End If
                End With
            Next
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixProcessHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Sub FixRegistryHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim sData As String, i As Long
    Dim lType As REG_VALUE_TYPE
    Dim aData() As String
    Dim bDouble As Boolean
    Dim sDelim As String
    
    'Note: REG_RESTORE_SAME - is a default type if it was not specified in the argument of 'AddRegToFix'
    
    'If Result.CureType And REGISTRY_BASED Then
        If AryPtr(result.Reg) Then
            For i = 0 To UBound(result.Reg)
                With result.Reg(i)
                    
                    If (.ActionType And USE_FEATURE_DISABLE) And g_bDelmodeDisabling Then
                        Exit Sub
                    End If
                    
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
                    
                    If .ActionType And RESTORE_VALUE Then

                        Reg.DelVal .Hive, .Key, .Param, .Redirected
                        
                        Select Case .ParamType
                        
                        Case REG_RESTORE_SZ
                            Reg.SetStringVal .Hive, .Key, .Param, CStr(.DefaultData), .Redirected
                        
                        Case REG_RESTORE_EXPAND_SZ
                            Reg.SetExpandStringVal .Hive, .Key, .Param, CStr(.DefaultData), .Redirected
                        
                        Case REG_RESTORE_BINARY
                            Reg.SetBinaryVal .Hive, .Key, .Param, HexStringToArray(CStr(.DefaultData)), .Redirected
                        
                        Case REG_RESTORE_DWORD
                            Reg.SetDwordVal .Hive, .Key, .Param, CLng(.DefaultData), .Redirected
                        
                        Case REG_RESTORE_QWORD
                            Reg.SetQwordVal .Hive, .Key, .Param, CLng(.DefaultData), .Redirected
                        
                        'Case REG_RESTORE_LINK
                        
                        Case REG_RESTORE_MULTI_SZ
                            aData = SplitSafe(CStr(.DefaultData), vbNullChar)
                            Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                        
                        End Select
                    End If
                    
                    If .ActionType And REMOVE_VALUE Then
                    
                        Reg.DelVal .Hive, .Key, .Param, .Redirected
                    End If
                
                    If .ActionType And REMOVE_KEY Then
                    
                        Reg.DelKey .Hive, .Key, .Redirected
                    End If
                    
                    If .ActionType And REMOVE_KEY_IF_NO_VALUES Then
                    
                        If Not Reg.KeyHasValues(.Hive, .Key, .Redirected) Then
                            Reg.DelKey .Hive, .Key, .Redirected
                        End If
                    End If
                    
                    If .ActionType And RESTORE_VALUE_INI Then
                    
                        IniSetString .IniFile, .Key, .Param, CStr(.DefaultData)
                    End If
                        
                    If .ActionType And REMOVE_VALUE_INI Then
                        
                        IniRemoveString .IniFile, .Key, .Param
                    End If
                    
                    If .ActionType And APPEND_VALUE_NO_DOUBLE Then
                    
                        'check if value already contains data planned to be written
                        bDouble = False
                        sDelim = ""
                        If Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then
                        
                            sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected, True)) 'true - do not expand
                            
                            If .ParamType = REG_RESTORE_MULTI_SZ Then
                                sDelim = vbNullChar
                            ElseIf .TrimDelimiter <> "" Then
                                sDelim = .TrimDelimiter
                            End If
                            If sDelim <> "" Then
                                bDouble = inArraySerialized(CStr(.DefaultData), sData, sDelim, , , vbTextCompare)
                            End If
                        End If
                        
                        If Not bDouble Then
                            'adding data to the beginning of the value
                            If Len(sData) = 0 Or Len(TrimEx(sData, sDelim)) = 0 Then
                                sData = .DefaultData
                            Else
                                sData = .DefaultData & sDelim & sData
                            End If
                            
                            Select Case .ParamType
                            
                            Case REG_RESTORE_MULTI_SZ
                                aData = SplitSafe(sData, vbNullChar)
                                Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                                
                            Case REG_RESTORE_SZ
                                Reg.SetStringVal .Hive, .Key, .Param, sData, .Redirected
                                
                            Case REG_RESTORE_EXPAND_SZ
                                Reg.SetExpandStringVal .Hive, .Key, .Param, sData, .Redirected
                                
                            End Select
                        End If
                    End If
                    
                    If .ActionType And REPLACE_VALUE Then
                    
                        If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                    
                        sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                        
                        If .ActionType And TRIM_VALUE Then 'if further 'trim' is planned to be in action
                            'replace by exact value
                            sData = Replace$(.TrimDelimiter & sData & .TrimDelimiter, _
                                .TrimDelimiter & .ReplaceDataWhat & .TrimDelimiter, _
                                .TrimDelimiter & .ReplaceDataInto & .TrimDelimiter, 1, 1, vbBinaryCompare) 'restrict to maximum 1 replacing
                        Else
                            'replace with part of value
                            sData = Replace$(sData, .ReplaceDataWhat, .ReplaceDataInto, 1, 1, vbTextCompare) 'restrict to maximum 1 replacing
                        End If
                        
                        lType = Reg.GetValueType(.Hive, .Key, .Param, .Redirected)
                        
                        Select Case lType
                        
                        Case REG_SZ
                            Reg.SetStringVal .Hive, .Key, .Param, sData, .Redirected
                                
                        Case REG_EXPAND_SZ
                            Reg.SetExpandStringVal .Hive, .Key, .Param, sData, .Redirected
                        
                        Case REG_MULTI_SZ
                            aData = SplitSafe(sData, vbNullChar)
                            Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                            
                        End Select
                    End If
                    
                    If .ActionType And TRIM_VALUE Then
                        
                        If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                        
                        sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                        
                        sData = TrimEx(sData, .TrimDelimiter)
                        
                        sData = Replace$(sData, .TrimDelimiter & .TrimDelimiter, .TrimDelimiter) '2 delims -> 1 delim
                        
                        lType = Reg.GetValueType(.Hive, .Key, .Param, .Redirected)
                        
                        Select Case lType
                        
                        Case REG_SZ
                            Reg.SetStringVal .Hive, .Key, .Param, sData, .Redirected
                                
                        Case REG_EXPAND_SZ
                            Reg.SetExpandStringVal .Hive, .Key, .Param, sData, .Redirected
                        
                        Case REG_MULTI_SZ
                            aData = SplitSafe(sData, vbNullChar)
                            Reg.SetMultiSZVal .Hive, .Key, .Param, aData(), .Redirected
                        
                        End Select
                    End If
                    
                    If .ActionType And REMOVE_VALUE_IF_EMPTY Then
                        
                        If Not Reg.ValueExists(.Hive, .Key, .Param, .Redirected) Then Exit Sub
                        
                        sData = CStr(Reg.GetData(.Hive, .Key, .Param, .Redirected))
                        If Len(sData) = 0 Then
                            Reg.DelVal .Hive, .Key, .Param, .Redirected
                        End If
                    End If
                    
                    If .ActionType And RESTORE_KEY_PERMISSIONS Then
                        RegKeyResetDACL .Hive, .Key, .Redirected, False
                    End If
                    
                    If .ActionType And RESTORE_KEY_PERMISSIONS_RECURSE Then
                        RegKeyResetDACL .Hive, .Key, .Redirected, True
                    End If

                End With
            Next
        End If
    'End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixRegistryHandler", result.HitLineW
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
        rType = REG_RESTORE_MULTI_SZ
        
    Case REG_ResourceList
        bRequiredDefault = True
        
    Case REG_FullResourceDescriptor
        bRequiredDefault = True
        
    Case REG_ResourceRequirementsList
        bRequiredDefault = True
        
    Case REG_QWORD
        rType = REG_RESTORE_QWORD
        
    Case REG_QWORD_LITTLE_ENDIAN
        rType = REG_RESTORE_QWORD
    
    Case Else
        bRequiredDefault = True
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

Public Sub FixFileHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    If result.CureType And FILE_BASED Then
        If AryPtr(result.File) Then
            For i = 0 To UBound(result.File)
                With result.File(i)
                    
                    If (.ActionType And USE_FEATURE_DISABLE) And g_bDelmodeDisabling Then
                        Exit Sub
                    End If
                
                    If .ActionType And UNREG_DLL Then
                        If Not IsMicrosoftFile(.Path, True) Or result.ForceMicrosoft Then
                            Reg.UnRegisterDll .Path
                        End If
                    End If
                    
                    If .ActionType And REMOVE_FILE Then
                        If FileExists(.Path) Then
                            DeleteFileWEx StrPtr(.Path), result.ForceMicrosoft
                        End If
                    End If
                    
                    If .ActionType And REMOVE_FOLDER Then
                        If FolderExists(.Path) Then
                            DeleteFolderForce .Path, result.ForceMicrosoft
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
                    
                    If .ActionType And CREATE_FOLDER Then
                        MkDirW .Path
                    End If
                End With
            Next
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixFileHandler", result.HitLineW
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

Public Sub FixServiceHandler(result As SCAN_RESULT)
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long, k As Long
    Dim aService() As String
    Dim aDepend() As String
    Dim bFixReg As Boolean
    
    If result.CureType And SERVICE_BASED Then
        If AryPtr(result.Service) Then
            For i = 0 To UBound(result.Service)
                With result.Service(i)
                
                    If (.ActionType And USE_FEATURE_DISABLE) And g_bDelmodeDisabling And (.ActionType And DELETE_SERVICE) Then
                        .ActionType = DISABLE_SERVICE
                    End If
                
                    If .ActionType And DELETE_SERVICE Then
                    
                        SetServiceStartMode .ServiceName, SERVICE_MODE_DISABLED
                        StopService .ServiceName
                        SetServiceStartMode .ServiceName, SERVICE_MODE_DISABLED
                        DeleteNTService .ServiceName, , result.ForceMicrosoft
                        
                        'Remove dependency
                        For j = 1 To Reg.EnumSubKeysToArray(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services", aService())
                            
                            aDepend = Reg.GetMultiSZ(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & aService(j), "DependOnService")
                            
                            If AryItems(aDepend) Then
                                For k = 0 To UBound(aDepend)
                                    If StrComp(aDepend(k), .ServiceName, 1) = 0 Then
                                        BackupKey result, HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\" & aService(j), "DependOnService"
                                        
                                        AddRegToFix result.Reg, REPLACE_VALUE Or TRIM_VALUE Or REMOVE_VALUE_IF_EMPTY, _
                                            HKLM, "System\CurrentControlSet\Services\" & aService(j), "DependOnService", , , REG_RESTORE_MULTI_SZ, _
                                            .ServiceName, "", vbNullChar
                                        
                                        result.CureType = result.CureType Or REGISTRY_BASED
                                        bFixReg = True
                                    End If
                                Next
                            End If
                        Next
                        If bFixReg Then
                            FixRegistryHandler result
                        End If
                    End If
                    
                    If .ActionType And DISABLE_SERVICE Then
                    
                        SetServiceStartMode .ServiceName, SERVICE_MODE_DISABLED
                    End If
                    
                    If .ActionType And RESTORE_SERVICE Then
                        '// TODO
                    End If
                    
                    If .ActionType And ENABLE_SERVICE Then
                        
                        SetServiceStartMode .ServiceName, SERVICE_MODE_AUTOMATIC
                    End If

                End With
            Next
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "FixServiceHandler", result.HitLineW
    If inIDE Then Stop: Resume Next
End Sub

Public Function CheckIntegrityHJT() As Boolean
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "CheckIntegrityHJT - Begin"
    'Checking consistency of HiJackThis.exe
    Dim SignResult As SignResult_TYPE
    Dim dModif As Date
    CheckIntegrityHJT = True
    If Not inIDE Then
        If OSver.IsWindows7OrGreater Then 'because my signature hash is SHA2
            If Not (OSver.SPVer < 1 And (OSver.MajorMinor = 6.1)) Then
                'ensure EDS subsystem is working correctly
                If isEDS_Work() Then
                    SignVerify AppPath(True), SV_PreferInternalSign, SignResult
                    If Not IsDragokasSign(SignResult) Then
                        'not a developer machine ?
                        dModif = GetFileDate(AppPath(True), DATE_MODIFIED)
                        If (GetDateAtMidnight(dModif) <> GetDateAtMidnight(Now()) And InStr(AppPath(), "_AVZ") = 0) Then
                            If (DateDiff("n", Now(), dModif) > 10) Or (DateDiff("n", Now(), dModif) < 0) Then
                                'Warning! Integrity of HiJackThis program is corrupted. Perhaps, file is patched or infected by file virus.
                                ErrReport = ErrReport & vbCrLf & Translate(1023) & vbCrLf
                                CheckIntegrityHJT = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    AppendErrorLogCustom "CheckIntegrityHJT - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CheckIntegrityHJT"
    If inIDE Then Stop: Resume Next
End Function

'ensure EDS subsystem is working correctly
Public Function isEDS_Work(Optional sReturnMsg As String) As Boolean
    Static bWork As Boolean
    Static bInit As Boolean
    Static sMsg As String
    Dim SignResult As SignResult_TYPE
    
    If Not bInit Then
        bInit = True
        SignVerify BuildPath(sWinDir, "system32\ntdll.dll"), SV_LightCheck Or SV_SelfTest, SignResult
        If IsMicrosoftCertHash(SignResult.HashRootCert) Then
            bWork = True
        End If
        sMsg = SignResult.ShortMessage & " (" & SignResult.FullMessage & ")"
    End If
    isEDS_Work = bWork
    sReturnMsg = sMsg
End Function

Public Function SetTaskBarProgressValue(frm As Form, ByVal Value As Single) As Boolean
    If Value < 0 Or Value > 1 Then Exit Function
    If Not (TaskBar Is Nothing) Then
        If Value = 0 Then
            TaskBar.SetProgressState g_HwndMain, TBPF_NOPROGRESS
        Else
            TaskBar.SetProgressValue frm.hwnd, CCur(Value * 10000), CCur(10000)
        End If
    End If
End Function

'// compare self version with installed one and update it
Public Function CheckInstalledVersionHJT() As Boolean
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "InstallUpdatedHJT - Begin"
    
    Dim sInstVer As String
    Dim HJT_Location As String
    HJT_Location = BuildPath(GetInstDir(), AppExeName(True))
    
    'if self
    If StrComp(AppPath(True), HJT_Location, 1) = 0 Then Exit Function
    
    'Installed version is present?
    If frmMain.chkConfigStartupScan.Value = 1 Then
        'compare versions
        sInstVer = GetFilePropVersion(HJT_Location)
        
        If ConvertVersionToNumber(sInstVer) < ConvertVersionToNumber(AppVerString) Then
            'The version of HiJackThis you launched is newer than installed one. Update it?
            If MsgBoxW(Translate(1402), vbQuestion Or vbYesNo) = vbYes Then
                'Replace 'Program Files' version with me
                CheckInstalledVersionHJT = InstallHJT(False)
            End If
        End If
    End If
    
    AppendErrorLogCustom "InstallUpdatedHJT - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "InstallUpdatedHJT"
    If inIDE Then Stop: Resume Next
End Function

Public Function InstallHJT(Optional bAskToCreateDesktopShortcut As Boolean, Optional bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim HJT_Location As String
    Dim HJT_LocationDir As String
    Dim sScanToolsDir As String
    Dim sScanToolsDirDest As String
    Dim bInstInPlace As Boolean
    Dim hFile As Long
    Dim aEXE() As String
    Dim i As Long
    
    AppendErrorLogCustom "InstallHJT - Begin"
    
    InstallHJT = True
    
    HJT_LocationDir = GetInstDir()
    HJT_Location = BuildPath(HJT_LocationDir, AppExeName(True))
    
    sScanToolsDir = BuildPath(AppPath(), "tools\Scan")
    sScanToolsDirDest = BuildPath(HJT_LocationDir, "tools\Scan")
    
    If StrComp(HJT_LocationDir, AppPath(), 1) = 0 Then
        bInstInPlace = True
    End If
    
    If Not bInstInPlace Then
        If Not MkDirW(HJT_LocationDir) Then
            MsgBoxW "Installation failed. Cannot create folder: " & HJT_LocationDir, vbCritical
            InstallHJT = False
            Exit Function
        End If
        'Copy exe to Program Files dir
        If Not FileCopyW(AppPath(True), HJT_Location, True) Then
            'MsgBoxW "Error while installing HiJackThis to program files folder. Cannot copy. Error = " & Err.LastDllError, vbCritical
            MsgBoxW Translate(593) & " " & Err.LastDllError, vbCritical
            InstallHJT = False
            Exit Function
        End If
        
        If FolderExists(sScanToolsDir) Then
            MkDirW sScanToolsDirDest
            FileCopyW BuildPath(sScanToolsDir, "auto.exe"), BuildPath(sScanToolsDirDest, "auto.exe")
            FileCopyW BuildPath(sScanToolsDir, "auto64.exe"), BuildPath(sScanToolsDirDest, "auto64.exe")
            FileCopyW BuildPath(sScanToolsDir, "executed.exe"), BuildPath(sScanToolsDirDest, "executed.exe")
            FileCopyW BuildPath(sScanToolsDir, "lastactivity.exe"), BuildPath(sScanToolsDirDest, "lastactivity.exe")
            FileCopyW BuildPath(sScanToolsDir, "serwin.exe"), BuildPath(sScanToolsDirDest, "serwin.exe")
            FileCopyW BuildPath(sScanToolsDir, "sheduler.exe"), BuildPath(sScanToolsDirDest, "sheduler.exe")
        End If
    End If
    
    If (FileExists(BuildPath(AppPath(), "whitelists.txt"))) Then
    
        If Not bInstInPlace Then
            FileCopyW BuildPath(AppPath(), "whitelists.txt"), BuildPath(HJT_LocationDir, "whitelists.txt")
        End If
        
        'Add HJT and supporting tools to exclude
        If OpenW(BuildPath(HJT_LocationDir, "whitelists.txt"), FOR_READ_WRITE, hFile) Then
            
            aEXE = ListFiles(BuildPath(HJT_LocationDir, "tools"), ".exe", True)
            If (AryPtr(aEXE)) Then
                For i = 0 To UBound(aEXE)
                    PrintW hFile, aEXE(i)
                Next
            End If
            PrintW hFile, AppPath(True)
            CloseW hFile
        End If
    End If
    
    'create Control panel -> 'Uninstall programs' entry
    CreateUninstallKey True, HJT_Location
    
    'Shortcuts in Start Menu
    InstallHJT = CreateHJTShortcuts(HJT_Location)
    
    If Not HasCommandLineKey("noShortcuts") Then
        If bAskToCreateDesktopShortcut Then
            If bSilent Then
                CreateHJTShortcutDesktop HJT_Location
            Else
                'Installation is completed. Do you want to create shortcut in Desktop?
                If MsgBoxW(Translate(69), vbYesNo, "HiJackThis") = vbYes Then
                    CreateHJTShortcutDesktop HJT_Location
                End If
            End If
        End If
    End If
    
    AppendErrorLogCustom "InstallHJT - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "InstallHJT"
    If inIDE Then Stop: Resume Next
End Function

Public Function InstallAutorunHJT(Optional bSilent As Boolean, Optional lDelay As Long = 60) As Boolean
    On Error GoTo ErrorHandler
    
    Dim JobCommand As String
    Dim HJT_Location As String
    Dim iExitCode As Long
    Dim HJT_Command As String
    Dim Delay As String
    Dim sArguments As String
    Dim pos As Long
    
    Delay = CStr(lDelay)
    
    AppendErrorLogCustom "InstallAutorunHJT - Begin"
    
    HJT_Location = BuildPath(GetInstDir(), AppExeName(True))
    
'    If MsgBox("This will install HJT to 'Program Files' folder and set Windows for automatically run HJT scan at system startup." & _
'        vbCrLf & vbCrLf & "Continue?" & vbCrLf & vbCrLf & "Note: it is recommended that you add all safe items to ignore list, so " & _
'        "the results window will appear at system startup if only new item will be found.", vbYesNo Or vbQuestion) = vbNo Then
    If Not bSilent Then
        If MsgBoxW(Translate(66), vbYesNo Or vbQuestion) = vbNo Then
            gNotUserClick = True
            frmMain.chkConfigStartupScan.Value = 0
            gNotUserClick = False
            Exit Function
        End If
        'To increase system loading speed it is recommended to set a delay
        'before launching HiJackThis on user logon. Specify the delay (in seconds):
        Delay = InputBox(Translate(1403), "HiJackThis", "60")
        If Not IsNumeric(Delay) Then Delay = "60"
        If CLng(Delay) < 0 Then Delay = "60"
    End If
    
    If OSver.IsWindowsVistaOrGreater Then
        'check if 'Schedule' service is launched
        If Not RunScheduler_Service(True, Not bSilent, bSilent) Then
            Exit Function
        End If
    End If
    
    pos = InStr(1, g_sCommandLine, "/!")
    
    If pos = 0 Then
        sArguments = "/startupscan"
    Else
        sArguments = Mid$(g_sCommandLine, pos + 3)
    End If
    
    If InstallHJT(, HasCommandLineKey("noGUI")) Then
    
        If OSver.IsWindowsVistaOrGreater Then
        
'            'delay after system startup for 1 min.
'            JobCommand = "/create /tn ""HiJackThis Autostart Scan"" /SC ONSTART /DELAY 0001:00 /F /RL HIGHEST " & _
'                "/tr ""\""" & HJT_Location & "\"" /startupscan"""
'
'            If Proc.ProcessRun("schtasks.exe", JobCommand, , 0) Then
'                iExitCode = Proc.WaitForTerminate(, , , 15000)     'if ExitCode = 0, 15 sec for timeout
'                If ERROR_SUCCESS <> iExitCode Then
'                    Proc.ProcessClose , , True
'                    'MsgBoxW "Error while creating task. Error = " & iExitCode, vbCritical
'                    MsgBoxW Translate(594) & " " & iExitCode, vbCritical
'                Else
'                    InstallAutorunHJT = True
'                End If
'            End If
            
            If CreateTask("HiJackThis Autostart Scan", HJT_Location, sArguments, _
              "Automatically scan the system with HiJackThis at user logon", CLng(Delay)) Then
                InstallAutorunHJT = True
            Else
                'MsgBoxW "Error while creating task", vbCritical
                MsgBoxW Translate(594), vbCritical
            End If
        Else
            'XP-
            'to add to 'Run' registry key
            HJT_Command = """" & HJT_Location & """" & " " & sArguments
            InstallAutorunHJT = Reg.SetExpandStringVal(HKLM, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis Autostart Scan", HJT_Command)
        End If
    End If
    
    AppendErrorLogCustom "InstallAutorunHJT - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "InstallAutorunHJT"
    If inIDE Then Stop: Resume Next
End Function

Public Function RemoveAutorunHJT() As Boolean
    If OSver.IsWindowsVistaOrGreater Then
        RemoveAutorunHJT = KillTask2("\HiJackThis Autostart Scan")
    Else
        RemoveAutorunHJT = Reg.DelVal(HKLM, "Software\Microsoft\Windows\CurrentVersion\Run", "HiJackThis Autostart Scan")
    End If
End Function

Public Sub OpenAndSelectFile(sFile As String)
    Dim hRet As Long
    Dim pidl As Long
    Dim sFileIDL As String
    
    If OSver.MajorMinor >= 5.1 Then '(XP+)
    
        sFileIDL = sFile
        
        'commented because such pIDL doesn't work on SHOpenFolderAndSelectItems
'        If OSver.IsWin64 Then
'            If StrBeginWith(sFileIDL, sWinSysDir) Then
'                sFileIDL = Replace$(sFileIDL, sWinSysDir, sWinDir & "\sysnative", 1, 1, 1)
'            End If
'        End If
    
        pidl = ILCreateFromPath(StrPtr(sFileIDL))

        If pidl <> 0 Then
            hRet = SHOpenFolderAndSelectItems(pidl, 0, 0, 0)
            
            ILFree pidl
        End If
    End If
    
    If pidl = 0 Or hRet <> S_OK Then
        'alternate
        If OSver.IsWin64 Then 'fix for 64 bit
            Shell sWinDir & "\explorer.exe /select," & """" & sFile & """", vbNormalFocus
        Else
            Shell "explorer.exe /select," & """" & sFile & """", vbNormalFocus
        End If
    End If
End Sub

Public Function GetDateAtMidnight(dDate As Date) As Date
    GetDateAtMidnight = DateAdd("s", -Second(dDate), DateAdd("n", -Minute(dDate), DateAdd("h", -Hour(dDate), dDate)))
End Function

Public Sub HJT_SaveReport(Optional nTry As Long)
    On Error GoTo ErrorHandler:
    Dim idx&
    
    AppendErrorLogCustom "HJT_SaveReport - Begin"

    idx = 7
    
    If bAutoLog Then
        If Len(g_sLogFile) = 0 Then
            g_sLogFile = BuildPath(AppPath(), "HiJackThis.log")
        End If
    Else
        bGlobalDontFocusListBox = True
        'sLogFile = SaveFileDialog("Save logfile...", "Log files (*.log)|*.log|All files (*.*)|*.*", "HiJackThis.log")
        g_sLogFile = SaveFileDialog(Translate(1001), AppPath(), "HiJackThis.log", Translate(1002) & " (*.log)|*.log|" & Translate(1003) & " (*.*)|*.*")
        bGlobalDontFocusListBox = False
    End If
    
    idx = 8
    
    If 0 <> Len(g_sLogFile) Then
        
        idx = 11
        
        Dim b() As Byte
        
        b = CreateLogFile() '<<<<<< ------- preparing all text for log file
        
        idx = 13
        
        'in /silentautolog mode log handle is already opened
        
        If g_hLog <= 0 Then
            If Not OpenW(g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog, g_FileBackupFlag) Then

                If Not bAutoLogSilent Then 'not via AutoLogger
                    'try another name

                    g_sLogFile = Left$(g_sLogFile, Len(g_sLogFile) - 4) & "_2.log"

                    Call OpenW(g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog)
                End If
            End If
        End If
        
        If g_hLog <= 0 Then
            If bAutoLogSilent Then 'via AutoLogger
                Exit Sub
            Else
            
                If bAutoLog Then ' if user clicked 1-st button (and HJT on ReadOnly media) => try another folder
                
                    bGlobalDontFocusListBox = True
                    'sLogFile = SaveFileDialog("Save logfile...", "Log files (*.log)|*.log|All files (*.*)|*.*", "HiJackThis.log")
                    g_sLogFile = SaveFileDialog(Translate(1001), AppPath(), "HiJackThis.log", Translate(1002) & " (*.log)|*.log|" & Translate(1003) & " (*.*)|*.*")
                    bGlobalDontFocusListBox = False
                    
                    If 0 <> Len(g_sLogFile) Then
                        If Not OpenW(g_sLogFile, FOR_OVERWRITE_CREATE, g_hLog) Then    '2-nd try
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
        
        PutW g_hLog, 1&, VarPtr(b(0)), UBound(b) + 1, doAppend:=True
        
        Dim lret As Long
        Dim ov As OVERLAPPED
        ov.offset = 0
        ov.InternalHigh = 0
        ov.hEvent = 0
        
        If g_LogLocked Then
            lret = UnlockFileEx(g_hLog, 0&, 1& * 1024 * 1024, 0&, VarPtr(ov))
            
            If lret Then
                g_LogLocked = False
            Else
                Debug.Print "UnlockFileEx is failed with err = " & Err.LastDllError
            End If
        End If
        
        CloseW g_hLog, True: g_hLog = 0
        
        'Check the size of the log
        If 0 = FileLenW(g_sLogFile) Then
            If nTry <> 2 Then
                SleepNoLock 100
                DeleteFileWEx StrPtr(g_sLogFile), , True
                SleepNoLock 400
                HJT_SaveReport 2
                Exit Sub
            End If
        End If
        
        idx = 14
        
        If (Not bAutoLogSilent) Or inIDE Then OpenLogFile g_sLogFile
    End If
    
    AppendErrorLogCustom "HJT_SaveReport - End"
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "HJT_SaveReport", "Stady: ", idx
    If inIDE Then Stop: Resume Next
End Sub

' Opens log file in default editor / or notepad if editor is not assigned to the extension
' / or in explorer window with selection if all other methods failed
'
Public Sub OpenLogFile(ByVal sLogFile As String)
    If Not FileExists(sLogFile, , True) Then Exit Sub
    
    Dim bAssoc As Boolean
    Dim bFailed As Boolean
    Dim sClassID As String
    Dim sOpenCmd As String
    Dim sOpenProg As String
    
    sLogFile = PathX64(sLogFile)
    
    sClassID = Reg.GetString(HKEY_CLASSES_ROOT, GetExtensionName(sLogFile), "")
    
    If sClassID <> "" Then
        sOpenCmd = EnvironW(Reg.GetString(HKEY_CLASSES_ROOT, sClassID & "\shell\open\command", ""))
        
        SplitIntoPathAndArgs sOpenCmd, sOpenProg, , True
        
        If FileExists(sOpenProg) Then
            bAssoc = True
        End If
    End If

    If bAssoc Then
        If OSver.IsWindowsXPOrGreater Then
            If Proc.ProcessRunUnelevated2(BuildPath(sWinDir, "explorer.exe"), sLogFile) Then Exit Sub
        End If
    End If
    
    If bAssoc Then
        bFailed = ShellExecute(g_HwndMain, StrPtr("open"), StrPtr(sLogFile), 0&, 0&, 1) <= 32
    End If
    
    If Not bAssoc Or bFailed Then
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
End Sub

Public Sub HJT_Shutdown()   ' emergency exits the program due to exceeding the timeout limit
    
    '!!! HiJackThis was shut down due to exceeding the maximum allowed timeout: [] sec. !!! Report file will be incomplete!
    'Please, restart the program manually (not via Autologger).
    ErrReport = ErrReport & vbCrLf & Replace$(Translate(1027), "[]", Perf.MAX_TimeOut)
    
    Dim s$
    If g_hDebugLog <> 0 Then
        s = vbCrLf & vbCrLf & String(39, "=") & vbCrLf & "!!! WARNING !!! Timeout is detected !!!" & vbCrLf & String(39, "=") & vbCrLf & vbCrLf
        PutW g_hDebugLog, 1, StrPtr(s), LenB(s), True
    End If
    
    'CloseW hLog, True
    'DeleteFileW BuildPath(AppPath(), "HiJackThis.log")
    
    'SortSectionsOfResultList
    HJT_SaveReport
    
    If inIDE Then
        Unload frmMain
        Debug.Print "HJT_Shutdown is raised!"
        End
    Else
        ExitProcess 1001&
    End If
End Sub

Public Function WhiteListed(sFile As String, sWhiteListedPath As String, Optional bCheckFileNamePartOnly As Boolean) As Boolean
    'to check matching the file with the specified name and verify it by EDS

    If bHideMicrosoft And Not bIgnoreAllWhitelists Then
        If bCheckFileNamePartOnly Then
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

Public Function SplitByMultiDelims(ByVal sLine As String, bFirstMatchOnly As Boolean, s_out_UsedDelim As String, ParamArray Delims()) As String()
    On Error GoTo ErrorHandler
    Dim i As Long
    If Not bFirstMatchOnly Then
        'replace all delimiters by first one
        For i = 1 To UBound(Delims)
            sLine = Replace$(sLine, Delims(i), Delims(0))
        Next
        s_out_UsedDelim = Delims(0)
        SplitByMultiDelims = SplitSafe(sLine, CStr(Delims(0)))
    Else
        For i = 0 To UBound(Delims)
            'substitute each delimiter
            If InStr(sLine, Delims(i)) <> 0 Then
                SplitByMultiDelims = SplitSafe(sLine, CStr(Delims(i)))
                s_out_UsedDelim = Delims(i)
                Exit Function
            End If
        Next
        'if no delimiters found, set initial string
        Dim arr(0) As String
        arr(0) = sLine
        SplitByMultiDelims = arr
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "SplitByMultiDelims", sLine
    ReDim SplitByMultiDelims(0)
    If inIDE Then Stop: Resume Next
End Function

' adds some warning to the end of the log (before debugging info)
Public Sub AddWarning(sMsg As String)
    EndReport = EndReport & vbCrLf & "Warning: " & sMsg
End Sub

'"Welcome to HJT" / or "Below are the results..." depending on "txtNothing" label vision state, availability of lstResults records and progressbar.
Public Sub pvSetVisionForLabelResults()
    If isRanHJT_Scan Then
        frmMain.lblInfo(0).Visible = False
        If Not frmMain.shpBackground.Visible Then
            ResumeProgressbar
        End If
    Else
        If frmMain.txtNothing.Visible Or frmMain.lstResults.ListCount <> 0 Then
            frmMain.lblInfo(1).Visible = True
            frmMain.lblInfo(0).Visible = False
        Else
            frmMain.lblInfo(1).Visible = False
            frmMain.lblInfo(0).Visible = True
        End If
    End If
End Sub

Public Function LenSafe(var As Variant) As Long
    If IsMissing(var) Then
        LenSafe = 0
    Else
        LenSafe = Len(CStr(var))
    End If
End Function

Public Function LoadResString(idFrom As Long, Optional idTo As Long) As String
    If idTo = 0 Then
        LoadResString = LoadResData(idFrom, 6)
    Else
        Dim i&, s$
        For i = idFrom To idTo
            s = s & LoadResData(i, 6)
        Next
        LoadResString = s
    End If
End Function

Public Function ConvertDateToUSFormat(d As Date) As String 'DD.MM.YYYY HH:MM:SS -> YYYY/MM/DD HH:MM:SS (for sorting purposes)
    ConvertDateToUSFormat = Format(d, "yyyy\/mm\/dd hh:nn:ss", vbMonday)
End Function

'@ sCmdLine - in. full command line
'@ sBaseKey - in. key to search subkeys for
'@ aKey - out. array of subkeys
'@ aValue - out. array of values corresponding to subkeys
'ret - number of items in "aKey" and "aValue" arrays
'
'SubKey example: /autostart d:600
'
Public Function ParseSubCmdLine(sCmdLine As String, sBaseKey As String, aKey() As String, aValue() As String)
    On Error GoTo ErrorHandler
    
    Dim pos As Long, pd As Long, cnt As Long
    Dim ch As String, sSearch As String

    pos = InStr(1, sCmdLine, sBaseKey, 1)
    If pos <> 0 Then
        pos = pos + Len(sBaseKey) + 1
        Do
            ReDim Preserve aKey(cnt)
            ReDim Preserve aValue(cnt)
            ch = Mid$(sCmdLine, pos, 1)
            If (ch = "-" Or ch = "/") Then Exit Do
            pd = InStr(pos, sCmdLine, ":")
            If pd = 0 Then Exit Do
            aKey(cnt) = LTrim(Mid$(sCmdLine, pos, pd - pos))
            If (Mid$(sCmdLine, pd + 1, 1) = """") Then
                pos = InStr(pd + 2, sCmdLine, """")
            Else
                pos = InStr(pd + 1, sCmdLine, " ")
            End If
            If (pos = 0) Then
                aValue(cnt) = Mid$(sCmdLine, pd + 1)
            Else
                aValue(cnt) = Mid$(sCmdLine, pd + 1, pos - pd - 1)
                pos = pos + 1
                Do While Mid$(sCmdLine, pos, 1) = " "
                    pos = pos + 1
                Loop
            End If
            cnt = cnt + 1
        Loop While pos
    End If
    If (cnt > 0) Then
        ReDim Preserve aKey(cnt - 1)
        ReDim Preserve aValue(cnt - 1)
    Else
        Erase aKey
        Erase aValue
    End If
    ParseSubCmdLine = cnt
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ParseSubCmdLine", sCmdLine, "BaseKey", sBaseKey
    If inIDE Then Stop: Resume Next
End Function

'@ sCmdLine - in. full command line
'@ sKey - in. key to get value of
'@ sValue - out. value of the specified key (if found)
'ret - true, if specified key was found
'
'Key example: /instDir:"c:\temp"
'
Public Function ParseCmdLineKey(ByVal sCmdLine As String, sKey As String, sValue As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim pos As Long
    Dim ch As String, sSearch As String

    pos = InStr(1, sCmdLine, sKey, 1)
    If pos <> 0 Then
        sCmdLine = Mid$(sCmdLine, pos + Len(sKey) + 1)
        ch = Left$(sCmdLine, 1)
        If ch = """" Then
            pos = InStr(2, sCmdLine, """")
            If pos <> 0 Then
                sValue = Mid$(sCmdLine, 2, pos - 2)
            End If
        Else
            pos = InStr(1, sCmdLine, " ")
            If pos <> 0 Then
                sValue = Left$(sCmdLine, pos - 1)
            Else
                sValue = sCmdLine
            End If
        End If
        ParseCmdLineKey = True
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ParseCmdLineKey", sCmdLine, "BaseKey", sKey
    If inIDE Then Stop: Resume Next
End Function

Public Function GetInstDir() As String
    Dim sValue As String
    Dim sInstDir As String
    If ParseCmdLineKey(g_sCommandLine, "instDir", sValue) Then
        sInstDir = sValue
        sInstDir = GetLongPath(sInstDir)
        sInstDir = GetFullPath(sInstDir)
    Else
        sInstDir = BuildPath(PF_32, "HiJackThis Fork")
    End If
    GetInstDir = sInstDir
End Function
