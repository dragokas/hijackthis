VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHiveEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsHiveEnum.cls]

'
' Hives Enumerator class by Alex Dragokas
'

'1.10
'Added .SID property (it's empty for non-HKCU/HKU hives)

'1.9
'HE_SID_NO_VIRTUAL is added to default .Init options. Specify zero to enum everything.
'".bak" and "{" char. in SID is added as sign of Virtual user.

'1.8
'Added new optional argument to .Init method: "Flags":
' - Added HE_CONTROLSET_ALL flag - adds ability for auto-enumeration of all ControlSets under HKLM\SYSTEM\ControlSetXXX
'(key should begin with SYSTEM\CurrentControlSet to activate this trigger; this flag is NOT set by default)
' - Added HE_CONTROLSET_EXCLUDE_CURRENT flag - to exclude SYSTEM\CurrentControlSet key itself from HE_CONTROLSET_ALL enumeration.
' - Added HE_KEY_MUST_EXIST flag - to enum. only existent keys.

'1.7
'Added method .Clear - to remove all keys and reset enumerator state (.Init data is saved)

'1.6.
'Added new methods: Uniq_AddKey, Uniq_MoveNext, Uniq_Repeat, Uniq_ClearKeys. New property: .Value
'It enum unique keys, manually specified by Uniq_AddKey method.

'1.5.
'Added support when program launched via Local System context (see constant #Hide_HKCU_For_Local_System)
'Added HE_SID_NO_VIRTUAL flag - allow excluding virtual users from HKU\<SID> enumeration, like 'Sandboxie' users (currently, it's alpha-version feature)

'1.4.
'Fixed wrong enumeration if first item of supplied array does not meet specified criteria
'Added support of HKCR hives

'1.3.
'Added .Clone function - to clone the class and its state

'1.2.
'Added recognition of NetShowServices' SID on Win2k

' ------ DESCRIPTION -------

'This class is intended to replace several 'for each' statements while iterating:
'1. HKLM / HKCU / HKU + each sid
'2. WOW and not WOW keys
'3. ControlSetXXX keys
'
'Additionally (if AddKeys / AddKey option is used, and BOTH redirection modes is requested):
' - It automatically recognize 'Shared' or 'Usual' (not 'redirected') keys
'     and doesn't produce WOW iteration for it avoiding access to the same phisical key again in current request.
'
'Also it allow to filter unique keys, specified by user (see .Uniq_ methods).
'
Option Explicit

'true, to return none when user query for HKCU hive and program is launched via Local System context
#Const Hide_HKCU_For_Local_System = True

'Private Enum KEY_INFORMATION_CLASS
'    KeyBasicInformation = 0
'    KeyNodeInformation
'    KeyFullInformation
'    KeyNameInformation
'    KeyCachedInformation
'    KeyFlagsInformation
'    KeyVirtualizationInformation
'    KeyHandleTagsInformation
'    MaxKeyInfoClass
'End Enum
'
'Private Enum OBJECT_INFORMATION_CLASS
'    ObjectBasicInformation = 0
'    ObjectNameInformation
'    ObjectTypeInformation
'    ObjectAllTypesInformation
'    ObjectHandleInformation
'End Enum
    
Private Type USER_INFO
    SID     As String
    Name    As String
    Virtual As Boolean
End Type
   
Private Type PACKED_CELL
    'in
    la_Keys()         As String
    la_Placeholder()  As String

    'out
    ls_Key            As String
    ls_KeyAndHive     As String
    ls_Value          As String 'For Uniq methods only
    '    le_Hive         As ENUM_REG_HIVE <- state_Hive used
    ls_HiveName       As String
    ls_HiveNameAndSID As String
    ls_SID            As String
    ls_UserName       As String
    '    lb_Redirection  As Boolean <- state_Redir used
    ll_KeyIndex       As Long

    'cache
    la_User()         As USER_INFO
    ls_CurUser        As String
    ls_CurSID         As String
    lb_User_dimmed    As Boolean
    lb_KeysPredefined As Boolean
    lb_OS_32          As Boolean
    ls_SID_NetShowServices As String
    lb_LocalSysContext As Boolean
    aControlSet()     As String

    'State
    lb_HaveKeys       As Boolean
    lb_HavePlaceHolders As Boolean
    '    l_SeqInit       As Long
    '    l_SeqKey        As Long
    state_CS          As Long
    state_SID         As Long
    lb_FirstMove      As Boolean
    lb_Class_Init     As Boolean
    lb_NewKey         As Boolean
    lb_Shared         As Boolean
    lb_SidService     As Boolean
    lb_SidUser        As Boolean
    lb_SidDefault     As Boolean
    EOF               As Boolean
End Type

Private cell As PACKED_CELL

'in
Private req_Hive          As HE_HIVE
Private req_SID           As HE_SID
Private req_Redir         As HE_REDIRECTION
Private req_Flags         As HE_OPTIONAL_FLAGS

'cache
Private la_predefHive()   As ENUM_REG_HIVE

'State
Private state_Hive        As ENUM_REG_HIVE
Private state_Redir       As ENUM_REG_REDIRECTION

Private Type UNIQ_REG_KEY_DATA
    hHive                 As ENUM_REG_HIVE
    KeyName               As String
    ValueName             As String
    Redir                 As Boolean
    KeyIndex              As Long 'user defined data
End Type

Private Type UNIQ_DATA
    Index                 As Long
    KeyData()             As UNIQ_REG_KEY_DATA
    dictKey               As clsTrickHashTable
End Type

Private UNIQ As UNIQ_DATA

'Private Declare Function CompareObjectHandles Lib "Kernelbase.dll" (ByVal hFirstObjectHandle As Long, ByVal hSecondObjectHandle As Long) As Long
'Private Declare Function NtQueryObject Lib "ntdll.dll" (ByVal Handle As Long, ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, ObjectInformation As Any, ByVal ObjectInformationLength As Long, ReturnLength As Long) As Long
'Private Declare Function NtQueryKey Lib "ntdll.dll" (ByVal KeyHandle As Long, ByVal KeyInformationClass As KEY_INFORMATION_CLASS, KeyInformation As Any, ByVal Length As Long, ResultLength As Long) As Long
'-> moved to 'bas'

Private Declare Function CreateNamedPipeW Lib "kernel32" (ByVal lpName As Long, ByVal dwOpenMode&, ByVal dwPipeMode&, ByVal nMaxInstances&, ByVal nOutBufferSize&, ByVal nInBufferSize&, ByVal nDefaultTimeOut&, ByVal lpSecurityAttributes&) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile&, lpBuffer As Any, ByVal nNumberOfBytesToWrite&, lpNumberOfBytesWritten&, ByVal lpOverlapped&) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile&, lpBuffer As Any, ByVal nNumberOfBytesToRead&, lpNumberOfBytesRead&, ByVal lpOverlapped&) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe&, lpBuffer As Any, ByVal nBufferSize&, lpBytesRead&, lpTotalBytesAvail&, lpBytesLeftThisMessage&) As Long
Private Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hPipe&) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject&) As Long

Private Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF

'Scheme of the cycle:
' {
'   1. Keys (if supplied)
'   {
'     2. HKLM / ControlSets / HKCU / HKU... <SIDs>
'     {
'       3. REDIR_WOW (redirected) / REDIR_NO_WOW
'     }
'   }
' }

'// set iterator settings about requested Hives / SIDs for HKU (if defined one) / WOW modes
Public Sub Init( _
    Hive_requested As HE_HIVE, _
    Optional SID_requested As HE_SID = HE_SID_ALL Or HE_SID_NO_VIRTUAL, _
    Optional Redirection_requested As HE_REDIRECTION = HE_REDIR_BOTH, _
    Optional Flags As HE_OPTIONAL_FLAGS = 0&)
    
    If Not cell.lb_Class_Init Then
        cell.lb_Class_Init = True
        cell.lb_OS_32 = Not bIsWOW64
    End If
    
    req_Hive = Hive_requested
    req_SID = SID_requested
    req_Redir = Redirection_requested
    req_Flags = Flags
    
    cell.lb_LocalSysContext = (OSver.SID_CurrentProcess() = "S-1-5-18")
    
    Clear
    
    'HE_CONTROLSET_EXCLUDE_CURRENT flag should be used together with HE_CONTROLSET_ALL
    If (req_Flags And HE_CONTROLSET_EXCLUDE_CURRENT) Then
        If Not CBool(req_Flags And HE_CONTROLSET_ALL) Then
            req_Flags = req_Flags Or HE_CONTROLSET_ALL
        End If
    End If
    If (req_Flags And HE_CONTROLSET_ALL) Then
        FillControlSet
    End If
End Sub

Private Sub FillControlSet()
    Static bInit As Boolean
    
    Dim iTotal As Long
    Dim i As Long
    Dim sKey As String
    
    If Not bInit Then
        bInit = True
        ReDim cell.aControlSet(0)
        cell.aControlSet(0) = "SYSTEM\CurrentControlSet"
        iTotal = 1
        For i = 1 To 999
            sKey = "SYSTEM\ControlSet" & Right$("00" & CStr(i), 3)
            If Reg.KeyExists(HKLM, sKey) Then
                ReDim Preserve cell.aControlSet(iTotal)
                cell.aControlSet(iTotal) = "SYSTEM\" & sKey
                iTotal = iTotal + 1
            Else
                Exit For
            End If
        Next
        ReDim Preserve cell.aControlSet(iTotal - 1)
    End If
End Sub

'// Clear last session
' Initialization data is saved
Public Sub Clear()
    Erase cell.la_Keys
    cell.lb_HaveKeys = False
    cell.lb_HavePlaceHolders = False
    cell.lb_KeysPredefined = False
    cell.ll_KeyIndex = -1
    cell.state_CS = 0
    ResetState
End Sub

'// Move through the same settings again
Public Sub Repeat()
    ResetState
End Sub

'// reset the state of iterator
Private Sub ResetState()
    cell.lb_FirstMove = True
    state_Hive = 0
    state_Redir = [_REG_REDIRECTION_NOT_DEFINED]
    cell.lb_NewKey = True
    cell.EOF = False
End Sub

'// optional. Set array of keys to iterate through and check appropriate state of iterator settings
'// Should be called after .Init

'NOTE:
'Allowed keys format:
'
'SOFTWARE
'HKLM\SOFTWARE
'HKEY_LOCAL_MACHINE\SOFTWARE
'
'You can mix several formats in one array.
'
'In case, hive prefix is not defined, requested prefixes (Hive_requested) will be appended during iteration.
'Otherwise, only predefined in array hive will be used.

Public Sub AddKeys(aKeys() As String)
    If Not cell.lb_Class_Init Then
        MsgBoxW "Class should be initialized first with .Init !!!"
        Exit Sub
    End If
    If AryItems(aKeys) Then
        cell.la_Keys = aKeys
        'FillPredefinedKeys
        ReDim Preserve cell.la_Placeholder(UBound(cell.la_Keys))
        cell.lb_HaveKeys = True
        cell.ll_KeyIndex = LBound(cell.la_Keys)
    End If
End Sub

'// optional. The same as above AddKeys(), but for adding +1 key, and also you can use placeholder
'Note: Can be used several times
'Note: you can append own 'Placeholder' to the key. It mean, class will not use this placeholder internally in checkings,
'  but it will return 'key + placeholder' in .Key and .KeyAndHive properties.
'  You can use such placeholder further, e.g. for replacing {CLSID} with real CLSID that you was unable to get before you initialized this class.
'Specify Hive_Apply to apply key to specific hive only
Public Sub AddKey(sKey As String, Optional sAppendPlaceholder As String)
    If Not cell.lb_Class_Init Then
        MsgBoxW "Class should be initialized first with .Init !!!"
        Exit Sub
    End If
    If Not cell.lb_HaveKeys Then
        ReDim cell.la_Keys(0)
        ReDim cell.la_Placeholder(0)
        cell.lb_HaveKeys = True
    Else
        ReDim Preserve cell.la_Keys(LBound(cell.la_Keys) To UBound(cell.la_Keys) + 1)
        ReDim Preserve cell.la_Placeholder(LBound(cell.la_Placeholder) To UBound(cell.la_Placeholder) + 1)
    End If
    cell.la_Keys(UBound(cell.la_Keys)) = sKey
    If 0 <> Len(sAppendPlaceholder) Then
        cell.lb_HavePlaceHolders = True
        cell.la_Placeholder(UBound(cell.la_Placeholder)) = sAppendPlaceholder
    End If
    cell.ll_KeyIndex = LBound(cell.la_Keys)
    'FillPredefinedKeys -> moved to .MoveNext on 'First call' since .AddKey can be called multiple times
End Sub

'// Splitting Hive\Key (if defined one) into 2 parts. In such case 'Hive' will not be iterated.
Private Function FillPredefinedKeys()
    On Error GoTo ErrorHandler:
    
    Dim i As Long
    Dim pos As Long
    'checking for predefined hives
    ReDim la_predefHive(LBound(cell.la_Keys) To UBound(cell.la_Keys))
    For i = LBound(cell.la_Keys) To UBound(cell.la_Keys)
        pos = InStr(cell.la_Keys(i), "\")
        If pos <> 0 Then
            'extract hive part -> convert to handle
            la_predefHive(i) = Reg.GetHKey(Left$(cell.la_Keys(i), pos - 1))
            If la_predefHive(i) <> 0 Then
                'trim to key part only
                cell.la_Keys(i) = Mid$(cell.la_Keys(i), pos + 1)
            End If
        End If
    Next
    cell.lb_KeysPredefined = True
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsHiveEnum.FillPredefinedKeys"
    If inIDE Then Stop: Resume Next
End Function

'// Go to the next state / or key (if defined one)
'increases the state (properties of this class)
'returns true, if required state is suitable for the key (if defined one).
'if not defined, returns true until all states will be reached
Public Function MoveNext() As Boolean
    On Error GoTo ErrorHandler:

    If cell.lb_FirstMove Then
        cell.lb_FirstMove = False
        If cell.lb_HaveKeys Then
            If Not cell.lb_KeysPredefined Then
                FillPredefinedKeys
            End If
        End If
        If Not NextHive() Then 'get the first hive
            MoveNext = MoveNextKey()
            Exit Function
        End If
        If Not NextWOW() Then 'get the first redirection state
            MoveNext = MoveNextKey() 'just in case
            Exit Function
        End If
        If Not KeySuitableWOW() Then
            MoveNext = MoveNext()
            Exit Function
        End If
    Else
        If cell.EOF Then Exit Function
        If NextWOW() Then               'iterate most embedded cycle (WOW mode)
            If Not KeySuitableWOW() Then
                MoveNext = MoveNext()
                Exit Function
            End If
        Else
            If NextHive() Then          '+1 to middle cycle (Hive and Sid)
                If Not KeySuitableWOW() Then
                    MoveNext = MoveNext()
                    Exit Function
                End If
            Else                        'last hive -> +1 to root cycle (key)
                MoveNext = MoveNextKey()
                If Not MoveNext Then cell.EOF = True
                Exit Function
            End If
        End If
    End If
    
    MoveNext = True
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsHiveEnum.MoveNext"
    If inIDE Then Stop: Resume Next
End Function

Private Function MoveNextKey() As Boolean
    If Not cell.lb_HaveKeys Then
        Exit Function       'no keys -> end
    Else
        If cell.ll_KeyIndex < UBound(cell.la_Keys) Then 'next key
            cell.ll_KeyIndex = cell.ll_KeyIndex + 1
            ResetState
            MoveNextKey = MoveNext()
            Exit Function
        Else
            Exit Function   'last key -> end
        End If
    End If
End Function

'// Move WOW state
'returns false, if the state is not suitable against requested mask or 'End of state' is reached
Private Function NextWOW() As Boolean
    
    If req_Redir = 0 Then Exit Function '(incorrect value), just in case
    
    If state_Redir = [_REG_REDIRECTION_NOT_DEFINED] Then    '[Not defined] -> REG_REDIRECTED
        state_Redir = REG_REDIRECTED
        
        If Not CBool(req_Redir And HE_REDIR_WOW) Then
            NextWOW = NextWOW()
            Exit Function
        End If
        
        'OS is x32 and requested both WOW modes?
        If cell.lb_OS_32 And (req_Redir And HE_REDIR_BOTH) Then
            NextWOW = NextWOW()
            Exit Function
        End If
    
    ElseIf state_Redir = REG_REDIRECTED Then             'REG_REDIRECTED -> REG_NOTREDIRECTED
        state_Redir = REG_NOTREDIRECTED
        
        If Not CBool(req_Redir And HE_REDIR_NO_WOW) Then
            NextWOW = NextWOW()
            Exit Function
        End If
    
    ElseIf state_Redir = REG_NOTREDIRECTED Then                'REG_NOTREDIRECTED -> REG_REDIRECTED and returns 'false'
        state_Redir = REG_REDIRECTED
        
        If Not CBool(req_Redir And HE_REDIR_WOW) Then
            Call NextWOW
            Exit Function
        End If
        
        'OS is x32 and requested both WOW modes?
        If cell.lb_OS_32 And (req_Redir And HE_REDIR_BOTH) Then
            Call NextWOW
            Exit Function
        End If
        
        Exit Function
    End If
    
    NextWOW = True
    
End Function

'// Move Hive state
'returns false, if the 'hive' state is not suitable against requested mask or 'End of state' is reached
Private Function NextHive() As Boolean
    
    On Error GoTo ErrorHandler:
    
    Dim sSID_Key    As String
    Dim bControlSet As Boolean
    
    cell.lb_SidService = False
    cell.lb_SidUser = False
    cell.lb_SidDefault = False
    
    If state_Hive = 0 Then          'NOT DEFINED -> HKLM
        state_Hive = HKLM
        If Not CBool(req_Hive And HE_HIVE_HKLM) Then
            NextHive = NextHive()
            Exit Function
        End If
        If cell.lb_HaveKeys Then
            If la_predefHive(cell.ll_KeyIndex) <> 0 Then
                If (la_predefHive(cell.ll_KeyIndex) <> HKLM) And (la_predefHive(cell.ll_KeyIndex) <> HKCR) Then
                    NextHive = NextHive()
                    Exit Function
                End If
            End If
        End If
        If (req_Flags And HE_CONTROLSET_ALL) Then
            If PathBeginWith(cell.la_Keys(cell.ll_KeyIndex), "SYSTEM\CurrentControlSet") Then
                If (cell.state_CS = 0) And (req_Flags And HE_CONTROLSET_EXCLUDE_CURRENT) Then
                    cell.state_CS = 1
                End If
                If cell.state_CS > 0 Then
                    If cell.state_CS > UBound(cell.aControlSet) Then
                        NextHive = NextHive()
                        Exit Function
                    End If
                End If
                'state_CS = state_CS + 1 'moved to the end of sub
                bControlSet = True
                state_Hive = 0
            End If
        End If
        cell.ls_HiveName = "HKLM"
        cell.ls_UserName = "All users"
        
    ElseIf state_Hive = HKLM Then   'HKLM -> HKCU
        state_Hive = HKCU
        If cell.lb_LocalSysContext Then
            NextHive = NextHive()
            Exit Function
        End If
        If Not CBool(req_Hive And HE_HIVE_HKCU) Then
            NextHive = NextHive()
            Exit Function
        End If
        If cell.lb_HaveKeys Then
            If la_predefHive(cell.ll_KeyIndex) <> 0 Then
                If la_predefHive(cell.ll_KeyIndex) <> HKCU And (la_predefHive(cell.ll_KeyIndex) <> HKCR) Then
                    NextHive = NextHive()
                    Exit Function
                End If
            End If
        End If
        cell.ls_HiveName = "HKCU"
        cell.ls_UserName = cell.ls_CurUser
        cell.ls_SID = cell.ls_CurSID
        
    ElseIf state_Hive = HKCU Then   'HKCU -> HKU (.DEFAULT)
        
        If req_SID <> 0 Then
            state_Hive = HKU '!!! don't touch this line order !!!
            cell.state_SID = -4   '!!! don't touch this line order !!!
            cell.ls_HiveName = "HKU"
            
            If Not CBool(req_Hive And HE_HIVE_HKU) Then
                Exit Function
            End If
            If cell.lb_HaveKeys Then
                If la_predefHive(cell.ll_KeyIndex) <> 0 Then
                    If la_predefHive(cell.ll_KeyIndex) <> HKU Then
                        Exit Function
                    End If
                End If
            End If
            If Not CBool(req_SID And HE_SID_DEFAULT) Then
                NextHive = NextHive()
                Exit Function
            End If
            cell.ls_UserName = "Default user" '(HKU\S-1-5-18)
            sSID_Key = ".DEFAULT"
            cell.ls_SID = "S-1-5-18"
            cell.lb_SidDefault = True
        Else
            Exit Function
        End If
    ElseIf state_Hive = HKU Then    'HKU -> index += 1
            
        Select Case cell.state_SID
        
        Case -4, -3, -2 '.DEFAULT (-4) -> S-1-5-19 (-2) -> S-1-5-20 (-1)
        
            If Not CBool(req_SID And HE_SID_SERVICE) Then
                cell.state_SID = -1
                NextHive = NextHive()
                Exit Function
            End If
            cell.state_SID = cell.state_SID + 1
                        
            If OSver.MajorMinor = 5 And Len(cell.ls_SID_NetShowServices) <> 0 Then
                'do nothing
            Else
                'skip state '-3'
                If cell.state_SID = -3 Then cell.state_SID = cell.state_SID + 1
            End If
            
            Select Case cell.state_SID
            Case -3: cell.ls_SID = cell.ls_SID_NetShowServices: cell.ls_UserName = "NetShowServices"
                        'HKU\S-1-5-18 -> is a symlink to HKU\.DEFAULT
            Case -2: cell.ls_SID = "S-1-5-19": cell.ls_UserName = "Local service"
            Case -1: cell.ls_SID = "S-1-5-20": cell.ls_UserName = "Network service"
            End Select
            sSID_Key = cell.ls_SID
            cell.lb_SidService = True
            
        Case Is >= -1 'S-1-5-20 -> Other users' SID
            
            If Not CBool(req_SID And HE_SID_USER) Then
                Exit Function
            End If
            If cell.lb_User_dimmed Then
                If cell.state_SID < UBound(cell.la_User) Then
                    cell.state_SID = cell.state_SID + 1
                    If CBool(req_SID And HE_SID_NO_VIRTUAL) Then
                        If cell.la_User(cell.state_SID).Virtual Then
                            NextHive = NextHive()
                            Exit Function
                        End If
                    End If
                    cell.ls_SID = cell.la_User(cell.state_SID).SID: cell.ls_KeyAndHive = "HKU": cell.ls_UserName = cell.la_User(cell.state_SID).Name
                    sSID_Key = cell.ls_SID
                    cell.lb_SidUser = True
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End Select
    End If
    
    cell.ls_Key = ""
    cell.ls_HiveNameAndSID = cell.ls_HiveName
    If state_Hive = HKU Then 'append SID
        cell.ls_Key = sSID_Key
        cell.ls_HiveNameAndSID = cell.ls_HiveNameAndSID & "\" & sSID_Key
    End If
    If cell.lb_HaveKeys Then 'append Key from array
        If bControlSet Then
            cell.ls_Key = cell.aControlSet(cell.state_CS) & Mid$(cell.la_Keys(cell.ll_KeyIndex), 25&)
            cell.state_CS = cell.state_CS + 1
        Else
            'add SID subkey if supplied
            cell.ls_Key = IIf(Len(cell.ls_Key) <> 0, cell.ls_Key & "\", "") & cell.la_Keys(cell.ll_KeyIndex)
        End If
    End If
    If cell.lb_HaveKeys Then
        If (la_predefHive(cell.ll_KeyIndex) = HKCR) Then 'append merged symlink prefix after convertion HKCR -> to HKCU/HKLM
            cell.ls_Key = "Software\Classes" & IIf(cell.ls_Key = "", "", "\" & cell.ls_Key)
        End If
    End If
    
    If (req_Flags And HE_KEY_MUST_EXIST) Then
        If Not Reg.KeyExists(state_Hive, cell.ls_Key, CBool(state_Redir)) Then
            NextHive = NextHive()
            Exit Function
        End If
    End If
    
    cell.ls_KeyAndHive = cell.ls_HiveName & IIf(Len(cell.ls_Key) <> 0, "\" & cell.ls_Key, "")
    
    NextHive = True
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsHiveEnum.NextHive"
    If inIDE Then Stop: Resume Next
End Function

'// checks if key supports the 'redirection' state specified:
' if 2 conditions are fulfilled:
'   1) array of keys (or key) SUPPLIED
'   2) mask of BOTH redirection modes requested
' when iterator reaches 'WOW' state, this routine checks if current key has virtual type: 'Redirected' (not 'Shared' and not 'Usual' key),
' to avoid double iteration of the same phisical key. In such case, if virtual type is not 'Redirected key', function returns false.
'
'e.g. HKLM\Software\Classed\Appid with KEY_WOW64_64KEY
'and  HKLM\Software\Classed\Appid without this flag (in WOW mode), mean:
'     HKLM\Software\Classed\WOW6432Node\Appid points to the same phisical location.
'so, no need to iterate it twice.
'
'The same is actual for any 'Usual' key, like, HKCU\Environment, because in BOTH redirection modes (with KEY_WOW64_64KEY flag and without,
'  opened handle points to the same phisical location.

Private Function KeySuitableWOW() As Boolean
    
    Dim kvt As KEY_VIRTUAL_TYPE
    
    If cell.lb_HaveKeys Then
        
        kvt = Reg.GetKeyVirtualType(state_Hive, cell.ls_Key)
        
        cell.lb_Shared = CBool(kvt And KEY_VIRTUAL_SHARED)
        
        If req_Redir And HE_REDIR_DONT_IGNORE_SHARED Then 'force flag is given
            KeySuitableWOW = True
            Exit Function
        End If
        
        If req_Redir = HE_REDIR_BOTH Then
            If state_Redir = REG_REDIRECTED Then
                If 0 = (kvt And KEY_VIRTUAL_REDIRECTED) Then
                    'exclude second 'shared' key
                    KeySuitableWOW = False
                Else
                    KeySuitableWOW = True
                End If
            Else
                KeySuitableWOW = True 'current iteration is on the state 'not redirected' -> always 'true'
            End If
        Else
            KeySuitableWOW = True 'requested one redirection state only -> always 'true'
        End If
    Else
        KeySuitableWOW = True 'no keys -> always 'true'
    End If
End Function

Private Function PathBeginWith(sPath, sBeginPart As String) As Boolean
    If StrComp(Left$(sPath, Len(sBeginPart)), sBeginPart, 1) = 0 Then
        If Len(sPath) = Len(sBeginPart) Or Mid$(sPath, Len(sBeginPart) + 1, 1) = "\" Then PathBeginWith = True
    End If
End Function

Private Sub Class_Initialize()
    GetHKU_Sids
End Sub

Private Sub Class_Terminate()
    Erase cell.la_User
    Erase UNIQ.KeyData
    Set UNIQ.dictKey = Nothing
End Sub

'// getting other users' SID and map it to the corresponding username
Private Sub GetHKU_Sids()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "clsHiveEnum.GetHKU_Sids - Begin"
    
    Dim i&, k&, sUsername$, aTmpSID() As String
    Dim bVirtual As Boolean
    Dim bAccessible As Boolean
    
    cell.ls_CurUser = OSver.UserName
    k = -1
    
    aTmpSID = SplitSafe(Reg.EnumSubKeys(HKEY_USERS, vbNullString), "|")
    
    For i = 0 To UBound(aTmpSID)
        If aTmpSID(i) Like "S-#-#-#*" And (Not StrEndWith(aTmpSID(i), "_Classes")) And _
          aTmpSID(i) <> "S-1-5-18" And aTmpSID(i) <> "S-1-5-19" And aTmpSID(i) <> "S-1-5-20" Then
          
            bVirtual = False
            sUsername = MapSIDToUsername(aTmpSID(i))
            
            '// TODO: maybe, better validate it using ConvertStringSidToSid() ?
            
            If sUsername = "unknown" Then
                sUsername = sUsername & ": " & aTmpSID(i)
                bVirtual = True
            ElseIf StrEndWith(aTmpSID(i), ".bak") Then
                bVirtual = True
            ElseIf InStr(aTmpSID(i), "{") <> 0 Then 'AV Sandboxie
                bVirtual = True
            End If
            
            If OSver.MajorMinor = 5 And sUsername = "NetShowServices" Then 'Win2k
                cell.ls_SID_NetShowServices = aTmpSID(i)
            Else
                If StrComp(sUsername, cell.ls_CurUser, 1) <> 0 Then
                    
                    bAccessible = True
                    
                    If Not OSver.IsElevated Then
                        Dim hKey As Long
                        If ERROR_SUCCESS = RegOpenKeyEx(HKEY_USERS, StrPtr(aTmpSID(i)), 0&, KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS, hKey) Then
                            RegCloseKey hKey
                        Else
                            bAccessible = False
                        End If
                    End If
                    
                    If bAccessible Then
                        k = k + 1
                        ReDim Preserve cell.la_User(k)
                        cell.la_User(k).SID = aTmpSID(i)
                        cell.la_User(k).Name = sUsername
                        cell.la_User(k).Virtual = bVirtual
                        cell.lb_User_dimmed = True
                    End If
                Else
                    'current user
                    cell.ls_CurSID = aTmpSID(i)
                End If
            End If
        End If
    Next i
    
    AppendErrorLogCustom "clsHiveEnum.GetHKU_Sids - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsHiveEnum.GetHKU_Sids"
    If inIDE Then Stop: Resume Next
End Sub

'// get key part of the registry key of current iteration, e.g. "SOFTWARE\Microsoft"
Public Property Get Key() As String
    If cell.lb_HavePlaceHolders Then
        If 0 <> Len(cell.la_Placeholder(cell.ll_KeyIndex)) Then
            Key = cell.ls_Key & cell.la_Placeholder(cell.ll_KeyIndex)
        Else
            Key = cell.ls_Key
        End If
    Else
        Key = cell.ls_Key
    End If
End Property
'// get key part together with hive prefix of current iteration, e.g.: "HKLM\SOFTWARE\Microsoft"
Public Property Get KeyAndHive() As String
    If cell.lb_HavePlaceHolders Then
        If 0 <> Len(cell.la_Placeholder(cell.ll_KeyIndex)) Then
            KeyAndHive = cell.ls_KeyAndHive & cell.la_Placeholder(cell.ll_KeyIndex)
        Else
            KeyAndHive = cell.ls_KeyAndHive
        End If
    Else
        KeyAndHive = cell.ls_KeyAndHive
    End If
End Property
'// get value part of Uniq key
Public Property Get Value() As String
    Value = cell.ls_Value
End Property
'// get hive pseudo-handle of current iteration
Public Property Get Hive() As ENUM_REG_HIVE
    Hive = state_Hive
End Property
'// get short hive name of current iteration, e.g. "HKLM"
Public Property Get HiveName() As String
    HiveName = cell.ls_HiveName
End Property
'// get short hive name + SID (if available) of current iteration, e.g. "HKU\S-1-5-19", or "HKU\.DEFAULT"
Public Property Get HiveNameAndSID() As String
    HiveNameAndSID = cell.ls_HiveNameAndSID
End Property
'// get SID of current iteration, e.g. "S-1-5-19"
Public Property Get SID() As String
    SID = cell.ls_SID
End Property
'// get user name associated with a hive of current iteration
Public Property Get UserName() As String
    UserName = cell.ls_UserName
End Property
'// get redirection state of current iteration
Public Property Get Redirected() As Boolean
    Redirected = state_Redir
End Property
'// get index of keys' array passed to the class used in current iteration (or -1 if no array is passed)
Public Property Get KeyIndex() As Long
    KeyIndex = cell.ll_KeyIndex
End Property
'// say, whether key is 'shared' virtual type
Public Property Get SharedKey() As Boolean
    SharedKey = cell.lb_Shared
End Property
'// say, if current HKU hive's sid is 'S-1-5-19' or 'S-1-5-20'
Public Property Get IsSidService() As Boolean
    IsSidService = cell.lb_SidService
End Property
'// say, if current HKU hive's sid belongs to one of user's account (this not include HKCU's SID)
Public Property Get IsSidUser() As Boolean
    IsSidUser = cell.lb_SidUser
End Property
'// say, if current HKU hive's sid belongs to 'Default user' account
Public Property Get IsSidDefault() As Boolean
    IsSidDefault = cell.lb_SidDefault
End Property

'// Clone the state and all data of this class to another class
Public Function Clone(DestinationClass As clsHiveEnum, _
    Optional DoNotUseValue As Variant, _
    Optional DoNotUseName As String) As Boolean
    
    On Error GoTo ErrorHandler:
    
    Dim hPipe As Long
    Dim ff As Integer
    Dim sPipeName As String
    Dim b() As Byte
    Dim Bytes As Long
    Dim BytesRead As Long
    Dim bIncoming As Boolean
    Dim bPipeRequired As Boolean
    Dim i As Long
    
    bIncoming = Not IsMissing(DoNotUseValue)
    
    If bIncoming Then
        If DoNotUseName = "cell" Then bPipeRequired = True
    Else
        bPipeRequired = True
    End If
    
    If bPipeRequired Then
        sPipeName = "\\.\pipe\vbPipeHiveEnum"
        
        hPipe = CreateNamedPipeW(StrPtr(sPipeName), 3, 0, 255, -1, -1, 0, 0)
        
        If hPipe = INVALID_HANDLE_VALUE Then Exit Function
        
        ff = FreeFile()
        Open sPipeName For Binary As #ff
    End If
    
    If bIncoming Then
        'incoming
        Select Case DoNotUseName
        Case "cell"
            'unpack B() to -> cell
            b = DoNotUseValue
            WriteFile hPipe, b(0), UBound(b) + 1, Bytes, 0
            If Bytes = UBound(b) + 1 Then
                Get #ff, 1, cell
                Clone = True
            End If
        'Otherwise, unpacking ENUMs
        Case "req_Hive"
            req_Hive = DoNotUseValue
            Clone = True
        Case "req_SID"
            req_SID = DoNotUseValue
            Clone = True
        Case "req_Redir"
            req_Redir = DoNotUseValue
            Clone = True
        Case "req_Flags"
            req_Flags = DoNotUseValue
            Clone = True
        Case "la_predefHive"
            la_predefHive = DoNotUseValue
            Clone = True
        Case "state_Hive"
            state_Hive = DoNotUseValue
            Clone = True
        Case "state_Redir"
            state_Redir = DoNotUseValue
            Clone = True
'        Case "la_HiveApplied()"
'            ReDim Preserve la_HiveApplied(DoNotUseIndex)
'            la_HiveApplied(DoNotUseIndex) = DoNotUseValue
'            Clone = True
        End Select
    Else
        'outcoming
        'pack cell to -> B()
        Put #ff, 1, cell
        PeekNamedPipe hPipe, ByVal 0&, 0, ByVal 0&, Bytes, 0
        If Bytes > 0 Then
            ReDim Preserve b(0 To Bytes - 1)
            ReadFile hPipe, b(0), Bytes, BytesRead, ByVal 0&
        End If
    End If
    
    'close pipe before sending
    If ff <> 0 Then Close #ff
    
    If hPipe <> 0 Then
        DisconnectNamedPipe hPipe
        CloseHandle hPipe
    End If
    
    If (Not bIncoming) And Bytes > 0 And Bytes = BytesRead Then
        If (DestinationClass Is Nothing) Then
            Set DestinationClass = New clsHiveEnum
        End If
        Clone = DestinationClass.Clone(Me, b, "cell")
        Clone = Clone And DestinationClass.Clone(Me, req_Hive, "req_Hive")
        Clone = Clone And DestinationClass.Clone(Me, req_SID, "req_SID")
        Clone = Clone And DestinationClass.Clone(Me, req_Redir, "req_Redir")
        Clone = Clone And DestinationClass.Clone(Me, req_Flags, "req_Flags")
        Clone = Clone And DestinationClass.Clone(Me, la_predefHive, "la_predefHive")
        Clone = Clone And DestinationClass.Clone(Me, state_Hive, "state_Hive")
        Clone = Clone And DestinationClass.Clone(Me, state_Redir, "state_Redir")
        
'        If cell.lb_HaveKeys Then
'            For i = 0 To UBound(la_HiveApplied)
'                Clone = Clone And DestinationClass.Clone(Me, la_HiveApplied(i), "la_HiveApplied()", i)
'            Next
'        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsHiveEnum.Clone"
    If inIDE Then Stop: Resume Next
End Function


'Methods to work with unique keys, added by user
'It allow to filter duplicte keys (by its Hive / Key name / Value name / Redirection state)

'Currently, only these properties are supported:
' - Hive
' - HiveName
' - HiveNameAndSID
' - Key
' - KeyAndHive
' - Redirected
' - KeyIndex

Public Function Uniq_ClearKeys()
    With UNIQ
        Erase .KeyData
        .Index = 0
        If Not (.dictKey Is Nothing) Then .dictKey.RemoveAll
    End With
End Function

Public Function Uniq_Exists(hHive As ENUM_REG_HIVE, ByVal KeyName As String, Optional ValueName As String, Optional bUseWow64 As Boolean) As Boolean
    Dim s As String
    
    If UNIQ.dictKey Is Nothing Then Exit Function
    
    s = BuildPath(CStr(hHive), KeyName, ValueName, CStr(bUseWow64))
    
    Uniq_Exists = UNIQ.dictKey.Exists(s)
End Function

'Note: KeyIndex - is a user defined data

Public Function Uniq_AddKey(hHive As ENUM_REG_HIVE, ByVal KeyName As String, Optional ValueName As String, Optional bUseWow64 As Boolean, Optional KeyIndex As Long) As Boolean
    
    Dim s As String
    
    With UNIQ
        If .dictKey Is Nothing Then
            Set .dictKey = New clsTrickHashTable
            .dictKey.CompareMode = vbTextCompare
        End If
        
        s = BuildPath(CStr(hHive), KeyName, ValueName, CStr(bUseWow64))
        
        If Not .dictKey.Exists(s) Then
        
            If 0 = AryPtr(.KeyData) Then
                ReDim .KeyData(0)
            Else
                ReDim Preserve .KeyData(UBound(.KeyData) + 1)
            End If
            
            .dictKey.Add (s), UBound(.KeyData)
            
            With .KeyData(UBound(.KeyData))
                .hHive = hHive
                .KeyName = KeyName
                .ValueName = ValueName
                .Redir = bUseWow64
                .KeyIndex = KeyIndex
            End With
            
            Uniq_AddKey = True
        End If
    End With
    
End Function

Public Function Uniq_MoveNext() As Boolean
    With UNIQ
        If 0 = AryPtr(.KeyData) Then Exit Function
    
        If .Index <= UBound(.KeyData) Then
            
            With .KeyData(.Index)
                state_Redir = .Redir
                state_Hive = .hHive
                cell.ls_HiveName = Reg.GetShortHiveName(Reg.GetHiveNameByHandle(.hHive))
                cell.ls_Key = .KeyName
                cell.ls_Value = .ValueName
                cell.ls_KeyAndHive = BuildPath(cell.ls_HiveName, cell.ls_Key)
                cell.ll_KeyIndex = .KeyIndex
                
                If .hHive = HKU Then
                    If Len(.KeyName) > 0 Then
                        cell.ls_HiveNameAndSID = BuildPath(cell.ls_HiveName, GetRootPath(.KeyName))
                    Else
                        cell.ls_HiveNameAndSID = cell.ls_HiveName
                    End If
                Else
                    cell.ls_HiveNameAndSID = cell.ls_HiveName
                End If
            End With
            
            .Index = .Index + 1
            Uniq_MoveNext = True
        End If
    End With
End Function

Public Sub Uniq_Repeat()
    UNIQ.Index = 0
End Sub


'Public Sub TestClass()
'    Dim HE As clsHiveEnum
'    Set HE = New clsHiveEnum
'
'    Dim aKey(1) As String
'
'    aKey(0) = "HKLM\Software\Classes\AppID"
'    aKey(1) = "Software\Classes\CLSID"
'
'    HE.Init HE_HIVE_HKLM Or HE_HIVE_HKU, HE_SID_ALL, HE_REDIR_BOTH
'
'    HE.AddKeys aKey
'
'    Do While HE.MoveNext
'        Debug.Print " --------- "
'        Debug.Print "Hive handle: " & HE.Hive
'        Debug.Print "Hive name:   " & HE.HiveName
'        Debug.Print "Hive + key:  " & HE.KeyAndHive
'        Debug.Print "Key:         " & HE.Key
'        Debug.Print "Redirected:  " & HE.Redirected
'        Debug.Print "Array index: " & HE.KeyIndex
'        Debug.Print "User name:   " & HE.UserName
'    Loop
'
'    Set HE = Nothing
'End Sub

'// test staff
Public Sub PrintAll()
    
    If Not cell.lb_Class_Init Then Debug.Print "Class is not initialized with .Init()": Exit Sub
    
    Do While Me.MoveNext
        Debug.Print " --------- "
        Debug.Print "Hive handle: " & Me.Hive
        Debug.Print "Hive name:   " & Me.HiveName
        Debug.Print "Hive + key:  " & Me.KeyAndHive
        Debug.Print "Key:         " & Me.Key
        Debug.Print "Redirected:  " & Me.Redirected
        Debug.Print "Array index: " & Me.KeyIndex
        Debug.Print "User name:   " & Me.UserName
        Debug.Print "Sid system?  " & Me.IsSidService
        Debug.Print "Sid user?    " & Me.IsSidUser
        Debug.Print "Sid default? " & Me.IsSidDefault
        Debug.Print "Shared key?  " & Me.SharedKey
        
        'Stop
    Loop
    
End Sub
