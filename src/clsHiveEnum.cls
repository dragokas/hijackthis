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
'
' Hives Enumerator class by Alex Dragokas
'
' v.1.1 (20.11.2017)

'intended to replace several 'for each' statements while iterating:
'1. HKLM / HKCU / HKU + each sid
'2. WOW and not WOW keys
'
'Additionally (if AddKeys / AddKey option is used, and BOTH redirection modes requested):
' - Automatically recognize 'Shared' or 'Usual' (not 'redirected') keys
'     and doesn't produce WOW iteration for it avoiding access to the same phisical key again in current request.

Option Explicit

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

'in
Private req_Hive          As HE_HIVE
Private req_SID           As HE_SID
Private req_Redir         As HE_REDIRECTION

'out
Private ls_Key            As String
Private ls_KeyAndHive     As String
'Private le_Hive         As ENUM_REG_HIVE <- state_Hive used
Private ls_HiveName       As String
Private ls_HiveNameAndSID As String
Private ls_UserName       As String
'Private lb_Redirection  As Boolean <- state_Redir used
Private ll_KeyIndex       As Long

'cache
Private la_SID()          As String
Private la_User()         As String
Private ls_CurUser        As String
Private ls_CurSID         As String
Private lb_User_dimmed    As Boolean
Private la_predefHive()   As ENUM_REG_HIVE
Private lb_KeysPredefined As Boolean
Private lb_OS_32          As Boolean

'State
Private la_Keys()         As String
Private la_Placeholder()  As String
Private lb_HaveKeys       As Boolean
Private lb_HavePlaceHolders As Boolean
'Private l_SeqInit       As Long
'Private l_SeqKey        As Long
Private state_Hive        As ENUM_REG_HIVE
Private state_SID         As Long
Private state_Redir       As ENUM_REG_REDIRECTION
Private lb_FirstMove      As Boolean
Private lb_Class_Init     As Boolean
Private lb_NewKey         As Boolean
Private lb_Shared         As Boolean
Private lb_SidService     As Boolean
Private lb_SidUser        As Boolean
Private lb_SidDefault     As Boolean

'Private Declare Function CompareObjectHandles Lib "Kernelbase.dll" (ByVal hFirstObjectHandle As Long, ByVal hSecondObjectHandle As Long) As Long
'Private Declare Function NtQueryObject Lib "ntdll.dll" (ByVal Handle As Long, ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, ObjectInformation As Any, ByVal ObjectInformationLength As Long, ReturnLength As Long) As Long
'Private Declare Function NtQueryKey Lib "ntdll.dll" (ByVal KeyHandle As Long, ByVal KeyInformationClass As KEY_INFORMATION_CLASS, KeyInformation As Any, ByVal Length As Long, ResultLength As Long) As Long
'-> moved to 'bas'

'Scheme of the cycle:
' {
'   1. Keys (if supplied)
'   {
'     2. HKLM / HKCU / HKU...
'     {
'       3. REDIR_WOW (redirected) / REDIR_NO_WOW
'     }
'   }
' }

'// set iterator settings about requested Hives / SIDs for HKU (if defined one) / WOW modes
Public Sub Init( _
    Hive_requested As HE_HIVE, _
    Optional SID_requested As HE_SID = HE_SID_ALL, _
    Optional Redirection_requested As HE_REDIRECTION = HE_REDIR_BOTH)
    
    If Not lb_Class_Init Then
        lb_Class_Init = True
        lb_OS_32 = Not bIsWOW64
    End If
    
    req_Hive = Hive_requested
    req_SID = SID_requested
    req_Redir = Redirection_requested
    
    'clear last session
    Erase la_Keys
    lb_HaveKeys = False
    lb_HavePlaceHolders = False
    lb_KeysPredefined = False
    
    ll_KeyIndex = -1
    
    ResetState
End Sub

'// Move through the same settings again
Public Sub Repeat()
    ResetState
End Sub

'// reset the state of iterator
Private Sub ResetState()
    lb_FirstMove = True
    state_Hive = 0
    state_Redir = [_REG_REDIRECTION_NOT_DEFINED]
    lb_NewKey = True
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
    If Not lb_Class_Init Then
        MsgBoxW "Class should be initialized first with .Init !!!"
        Exit Sub
    End If
    If AryPtr(aKeys) Then
        la_Keys = aKeys
        'FillPredefinedKeys
        lb_HaveKeys = True
        ll_KeyIndex = LBound(la_Keys)
    End If
End Sub

'// optional. The same as above AddKeys(), but for adding +1 key, and also you can use placeholder
'Note: Can be used several times
'Note: you can append own 'Placeholder' to the key. It mean, class will not use this placeholder internally in checkings,
'  but it will return 'key + placeholder' in .Key and .KeyAndHive properties.
'  You can use such placeholder further, e.g. for replacing {CLSID} with real CLSID that you was unable to get before you initialized this class.
Public Sub AddKey(sKey As String, Optional sAppendPlaceholder As String)
    If Not lb_Class_Init Then
        MsgBoxW "Class should be initialized first with .Init !!!"
        Exit Sub
    End If
    If Not lb_HaveKeys Then
        ReDim la_Keys(0)
        ReDim la_Placeholder(0)
        lb_HaveKeys = True
    Else
        ReDim Preserve la_Keys(LBound(la_Keys) To UBound(la_Keys) + 1)
        ReDim Preserve la_Placeholder(LBound(la_Placeholder) To UBound(la_Placeholder) + 1)
    End If
    la_Keys(UBound(la_Keys)) = sKey
    If 0 <> Len(sAppendPlaceholder) Then
        lb_HavePlaceHolders = True
        la_Placeholder(UBound(la_Placeholder)) = sAppendPlaceholder
    End If
    ll_KeyIndex = LBound(la_Keys)
    'FillPredefinedKeys -> moved to .MoveNext on 'First call' since .AddKey can be called multiple times
End Sub

'// Splitting Hive\Key (if defined one) into 2 parts. In such case 'Hive' will not be iterated.
Private Function FillPredefinedKeys()
    On Error GoTo ErrorHandler:
    
    Dim i As Long
    Dim pos As Long
    'checking for predefined hives
    ReDim la_predefHive(LBound(la_Keys) To UBound(la_Keys))
    For i = LBound(la_Keys) To UBound(la_Keys)
        pos = InStr(la_Keys(i), "\")
        If pos <> 0 Then
            'extract hive part -> convert to handle
            la_predefHive(i) = Reg.GetHKey(Left$(la_Keys(i), pos - 1))
            If la_predefHive(i) <> 0 Then
                'trim to key part only
                la_Keys(i) = Mid$(la_Keys(i), pos + 1)
            End If
        End If
    Next
    lb_KeysPredefined = True
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FillPredefinedKeys"
    If inIDE Then Stop: Resume Next
End Function

'// Go to the next state / or key (if defined one)
'increases the state (properties of this class)
'returns true, if required state is suitable for the key (if defined one).
'if not defined, returns true until all states will be reached
Public Function MoveNext() As Boolean
    On Error GoTo ErrorHandler:

    If lb_FirstMove Then
        lb_FirstMove = False
        If lb_HaveKeys Then
            If Not lb_KeysPredefined Then
                FillPredefinedKeys
            End If
        End If
        If Not NextHive() Then 'get the first hive
            Exit Function 'just in case
        End If
        If Not NextWOW() Then 'get the first redirection state
            Exit Function 'just in case
        End If
        If Not KeySuitableWOW() Then
            MoveNext = MoveNext()
            Exit Function
        End If
    Else
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
                If Not lb_HaveKeys Then
                    Exit Function       'no keys -> end
                Else
                    If ll_KeyIndex < UBound(la_Keys) Then 'next key
                        ll_KeyIndex = ll_KeyIndex + 1
                        ResetState
                        MoveNext = MoveNext()
                        Exit Function
                    Else
                        Exit Function   'last key -> end
                    End If
                End If
            End If
        End If
    End If
    
    MoveNext = True
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "MoveNext"
    If inIDE Then Stop: Resume Next
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
        If lb_OS_32 And (req_Redir And HE_REDIR_BOTH) Then
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
        If lb_OS_32 And (req_Redir And HE_REDIR_BOTH) Then
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

    Dim sSID As String
    
    lb_SidService = False
    lb_SidUser = False
    lb_SidDefault = False

    If state_Hive = 0 Then          'NOT DEFINED -> HKLM
        state_Hive = HKLM
        If Not CBool(req_Hive And HE_HIVE_HKLM) Then
            NextHive = NextHive()
            Exit Function
        End If
        If lb_HaveKeys Then
            If la_predefHive(ll_KeyIndex) <> 0 Then
                If la_predefHive(ll_KeyIndex) <> HKLM Then
                    NextHive = NextHive()
                    Exit Function
                End If
            End If
        End If
        ls_HiveName = "HKLM"
        ls_UserName = "All users"
        
    ElseIf state_Hive = HKLM Then   'HKLM -> HKCU
        state_Hive = HKCU
        If Not CBool(req_Hive And HE_HIVE_HKCU) Then
            NextHive = NextHive()
            Exit Function
        End If
        If lb_HaveKeys Then
            If la_predefHive(ll_KeyIndex) <> 0 Then
                If la_predefHive(ll_KeyIndex) <> HKCU Then
                    NextHive = NextHive()
                    Exit Function
                End If
            End If
        End If
        ls_HiveName = "HKCU"
        ls_UserName = ls_CurUser
        
    ElseIf state_Hive = HKCU Then   'HKCU -> HKU (.DEFAULT)
        
        If req_SID <> 0 Then
            state_Hive = HKU '!!! don't touch this line order !!!
            state_SID = -4   '!!! don't touch this line order !!!
            ls_HiveName = "HKU"
            
            If Not CBool(req_Hive And HE_HIVE_HKU) Then
                Exit Function
            End If
            If lb_HaveKeys Then
                If la_predefHive(ll_KeyIndex) <> 0 Then
                    If la_predefHive(ll_KeyIndex) <> HKU Then
                        Exit Function
                    End If
                End If
            End If
            If Not CBool(req_SID And HE_SID_DEFAULT) Then
                NextHive = NextHive()
                Exit Function
            End If
            ls_UserName = "Default user" '(HKU\S-1-5-18)
            sSID = ".DEFAULT"
            lb_SidDefault = True
            
        Else
            Exit Function
        End If
    ElseIf state_Hive = HKU Then    'HKU -> index += 1
    
        Select Case state_SID
        
        Case -4, -3, -2 '.DEFAULT (-4) -> S-1-5-19 (-2) -> S-1-5-20 (-1)
        
            If Not CBool(req_SID And HE_SID_SERVICE) Then
                state_SID = -1
                NextHive = NextHive()
                Exit Function
            End If
            state_SID = state_SID + 1
            
            If state_SID = -3 Then state_SID = state_SID + 1
            
            Select Case state_SID
            Case -3: 'sSID = "S-1-5-18": ls_UserName = "System" 'removed, since HKU\S-1-5-18 -> is a symlink to HKU\.DEFAULT
            Case -2: sSID = "S-1-5-19": ls_UserName = "Local service"
            Case -1: sSID = "S-1-5-20": ls_UserName = "Network service"
            End Select
            lb_SidService = True
            
        Case Is >= -1 'S-1-5-20 -> Other users' SID
            
            If Not CBool(req_SID And HE_SID_USER) Then
                Exit Function
            End If
            If lb_User_dimmed Then
                If state_SID < UBound(la_SID) Then
                    state_SID = state_SID + 1
                    sSID = la_SID(state_SID): ls_KeyAndHive = "HKU": ls_UserName = la_User(state_SID)
                    lb_SidUser = True
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End Select
    End If
    
    ls_Key = ""
    ls_HiveNameAndSID = ls_HiveName
    If state_Hive = HKU Then 'append SID
        ls_Key = sSID
        ls_HiveNameAndSID = ls_HiveNameAndSID & "\" & sSID
    End If
    If lb_HaveKeys Then 'append Key from array
        ls_Key = IIf(Len(ls_Key) <> 0, ls_Key & "\", "") & la_Keys(ll_KeyIndex)
    End If
    ls_KeyAndHive = ls_HiveName & IIf(Len(ls_Key) <> 0, "\" & ls_Key, "")
    
    NextHive = True
    
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
    
    If lb_HaveKeys Then
        
        kvt = Reg.GetKeyVirtualType(state_Hive, ls_Key)
        
        lb_Shared = CBool(kvt And KEY_VIRTUAL_SHARED)
        
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

Private Sub Class_Initialize()
    GetHKU_Sids
End Sub

Private Sub Class_Terminate()
    Erase la_SID
    Erase la_User
End Sub

'// getting other users' SID and map it to the corresponding username
Private Sub GetHKU_Sids()
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetHKU_Sids - Begin"
    
    Dim i&, K&, sUsername$, aTmpSID() As String
    
    ls_CurUser = GetUser()
    K = -1
    
    aTmpSID = SplitSafe(Reg.EnumSubKeys(HKEY_USERS, vbNullString), "|")
    
    For i = 0 To UBound(aTmpSID)
        If aTmpSID(i) Like "S-#-#-#*" And (Not StrEndWith(aTmpSID(i), "_Classes")) And _
          aTmpSID(i) <> "S-1-5-18" And aTmpSID(i) <> "S-1-5-19" And aTmpSID(i) <> "S-1-5-20" Then
          
            sUsername = MapSIDToUsername(aTmpSID(i))
            
            If StrComp(sUsername, ls_CurUser, 1) <> 0 Then
                K = K + 1
                ReDim Preserve la_SID(K)
                ReDim Preserve la_User(K)
                la_SID(K) = aTmpSID(i)
                la_User(K) = sUsername
                lb_User_dimmed = True
            Else
                'current user
                ls_CurSID = aTmpSID(i)
            End If
        End If
    Next i
    
    AppendErrorLogCustom "GetHKU_Sids - End"
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetHKU_Sids"
    If inIDE Then Stop: Resume Next
End Sub

'// get key part of the registry key of current iteration, e.g. "SOFTWARE\Microsoft"
Public Property Get Key() As String
    If lb_HavePlaceHolders Then
        If 0 <> Len(la_Placeholder(ll_KeyIndex)) Then
            Key = ls_Key & la_Placeholder(ll_KeyIndex)
        Else
            Key = ls_Key
        End If
    Else
        Key = ls_Key
    End If
End Property
'// get key part together with hive prefix of current iteration, e.g.: "HKLM\SOFTWARE\Microsoft"
Public Property Get KeyAndHive() As String
    If lb_HavePlaceHolders Then
        If 0 <> Len(la_Placeholder(ll_KeyIndex)) Then
            KeyAndHive = ls_KeyAndHive & la_Placeholder(ll_KeyIndex)
        Else
            KeyAndHive = ls_KeyAndHive
        End If
    Else
        KeyAndHive = ls_KeyAndHive
    End If
End Property
'// get hive pseudo-handle of current iteration
Public Property Get Hive() As ENUM_REG_HIVE
    Hive = state_Hive
End Property
'// get short hive name of current iteration, e.g. "HKLM"
Public Property Get HiveName() As String
    HiveName = ls_HiveName
End Property
'// get short hive name + SID (if available) of current iteration, e.g. "HKU\S-1-5-19", or "HKU\.DEFAULT"
Public Property Get HiveNameAndSID() As String
    HiveNameAndSID = ls_HiveNameAndSID
End Property
'// get user name associated with a hive of current iteration
Public Property Get UserName() As String
    UserName = ls_UserName
End Property
'// get redirection state of current iteration
Public Property Get Redirected() As Boolean
    Redirected = state_Redir
End Property
'// get index of keys' array passed to the class used in current iteration (or -1 if no array is passed)
Public Property Get KeyIndex() As Long
    KeyIndex = ll_KeyIndex
End Property
'// say, whether key is 'shared' virtual type
Public Property Get SharedKey() As Boolean
    SharedKey = lb_Shared
End Property
'// say, if current HKU hive's sid is 'S-1-5-19' or 'S-1-5-20'
Public Property Get IsSidService() As Boolean
    IsSidService = lb_SidService
End Property
'// say, if current HKU hive's sid belongs to one of user's account (this not include HKCU's SID)
Public Property Get IsSidUser() As Boolean
    IsSidUser = lb_SidUser
End Property
'// say, if current HKU hive's sid belongs to 'Default user' account
Public Property Get IsSidDefault() As Boolean
    IsSidDefault = lb_SidDefault
End Property

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
    
    If Not lb_Class_Init Then Debug.Print "Class is not initialized with .Init()": Exit Sub
    
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
