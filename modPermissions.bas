Attribute VB_Name = "modPermissions"
'
' Reset Permissions Module by Alex Dragokas
'
' ver. 1.0
'
' This module is a part of HiJackThis project
'
Option Explicit

Private Enum ACCESS_MODE
    NOT_USED_ACCESS = 0
    GRANT_ACCESS
    SET_ACCESS
    DENY_ACCESS
    REVOKE_ACCESS
    SET_AUDIT_SUCCESS
    SET_AUDIT_FAILURE
End Enum

Private Enum TRUSTEE_FORM
    TRUSTEE_IS_SID = 0
    TRUSTEE_IS_NAME
    TRUSTEE_BAD_FORM
    TRUSTEE_IS_OBJECTS_AND_SID
    TRUSTEE_IS_OBJECTS_AND_NAME
End Enum

Private Enum TRUSTEE_TYPE
    TRUSTEE_IS_UNKNOWN = 0
    TRUSTEE_IS_USER
    TRUSTEE_IS_GROUP
    TRUSTEE_IS_DOMAIN
    TRUSTEE_IS_ALIAS
    TRUSTEE_IS_WELL_KNOWN_GROUP
    TRUSTEE_IS_DELETED
    TRUSTEE_IS_INVALID
    TRUSTEE_IS_COMPUTER
End Enum

Private Type TRUSTEE
    pMultipleTrustee As Long
    MultipleTrusteeOperation As Long
    TrusteeForm As TRUSTEE_FORM
    TrusteeType As TRUSTEE_TYPE
    ptstrName As Long
End Type

Private Type EXPLICIT_ACCESS
    grfAccessPermissions As Long
    grfAccessMode As ACCESS_MODE
    grfInheritance As Long
    tTrustee As TRUSTEE
End Type

Private Type ACE_HEADER
    AceType As Byte
    AceFlags As Byte
    AceSize As Integer
End Type

Private Type ACCESS_DENIED_ACE
    Header As ACE_HEADER
    Mask As Long 'ACCESS_MASK
    SidStart As Long
End Type

Private Type ACL_SIZE_INFORMATION
    AceCount As Long
    AclBytesInUse As Long
    AclBytesFree As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount  As Long
    LuidLowPart     As Long
    LuidHighPart    As Long
    Attributes      As Long
End Type

Private Type SID
    Revision As Byte
    SubAuthorityCount As Byte
    IdentifierAuthority(5) As Byte
    SubAuthority As Long
End Type

Private Enum ACL_INFORMATION_CLASS
    AclRevisionInformation = 1
    AclSizeInformation
End Enum

Private Enum SE_OBJECT_TYPE
    SE_UNKNOWN_OBJECT_TYPE = 0
    SE_FILE_OBJECT
    SE_SERVICE
    SE_PRINTER
    SE_REGISTRY_KEY
    SE_LMSHARE
    SE_KERNEL_OBJECT
    SE_WINDOW_OBJECT
    SE_DS_OBJECT
    SE_DS_OBJECT_ALL
    SE_PROVIDER_DEFINED_OBJECT
    SE_WMIGUID_OBJECT
    SE_REGISTRY_WOW64_32KEY
End Enum

Private Enum SECURITY_INFORMATION                       'required access - to query / to set info:
    ATTRIBUTE_SECURITY_INFORMATION = &H20&              'query: READ_CONTROL; set: WRITE_DAC
    BACKUP_SECURITY_INFORMATION = &H10000               'query: READ_CONTROL and ACCESS_SYSTEM_SECURITY; set: WRITE_DAC and WRITE_OWNER and ACCESS_SYSTEM_SECURITY
    DACL_SECURITY_INFORMATION = 4                       'query: READ_CONTROL; set: WRITE_DAC
    GROUP_SECURITY_INFORMATION = 2                      'query: READ_CONTROL; set: WRITE_OWNER
    LABEL_SECURITY_INFORMATION = 16                     'query: READ_CONTROL; set: WRITE_OWNER
    OWNER_SECURITY_INFORMATION = 1                      'query: READ_CONTROL; set: WRITE_OWNER
    PROTECTED_DACL_SECURITY_INFORMATION = &H80000000    'query: -; set: WRITE_DAC
    PROTECTED_SACL_SECURITY_INFORMATION = &H40000000    'query: -; set: ACCESS_SYSTEM_SECURITY
    SACL_SECURITY_INFORMATION = 8                       'query: ACCESS_SYSTEM_SECURITY; set: ACCESS_SYSTEM_SECURITY
    SCOPE_SECURITY_INFORMATION = &H40&                  'query: READ_CONTROL; set: ACCESS_SYSTEM_SECURITY
    UNPROTECTED_DACL_SECURITY_INFORMATION = &H20000000  'query: -; set: WRITE_DAC
    UNPROTECTED_SACL_SECURITY_INFORMATION = &H10000000  'query: -; set: ACCESS_SYSTEM_SECURITY
End Enum


Private Declare Sub GetNativeSystemInfo Lib "kernel32.dll" (ByVal lpSystemInfo As Long)
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueW" (ByVal lpSystemName As Long, ByVal lpName As Long, lpLuid As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long
Private Declare Function ConvertStringSidToSid Lib "advapi32.dll" Alias "ConvertStringSidToSidW" (ByVal StringSid As Long, pSid As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function CopySid Lib "advapi32.dll" (ByVal nDestinationSidLength As Long, ByVal pDestinationSid As Long, ByVal pSourceSid As Long) As Long
Private Declare Function GetLengthSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
Private Declare Function IsValidSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
'Private Declare Function RegDeleteKeyEx Lib "advapi32.dll" Alias "RegDeleteKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal samDesired As Long, ByVal Reserved As Long) As Long
Private Declare Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal Handle As Long, ByVal RequestedInformation As SECURITY_INFORMATION, ByVal pSecurityDescriptor As Long, ByVal nLength As Long, ByVal lpnLengthNeeded As Long) As Long
Private Declare Function MakeAbsoluteSD Lib "advapi32.dll" (ByVal pSelfRelativeSD As Long, ByVal pAbsoluteSD As Long, ByVal lpdwAbsoluteSDSize As Long, ByVal pDacl As Long, ByVal lpdwDaclSize As Long, ByVal pSacl As Long, ByVal lpdwSaclSize As Long, ByVal pOwner As Long, ByVal lpdwOwnerSize As Long, ByVal pPrimaryGroup As Long, ByVal lpdwPrimaryGroupSize As Long) As Long
Private Declare Function IsValidSecurityDescriptor Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long) As Long
Private Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclW" (ByVal cCountOfExplicitEntries As Long, ByVal pListOfExplicitEntries As Long, ByVal pOldAcl As Long, NewAcl As Long) As Long
Private Declare Function SetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As SECURITY_INFORMATION, ByVal psidOwner As Long, ByVal psidGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long) As Long
Private Declare Function SetNamedSecurityInfo Lib "advapi32.dll" Alias "SetNamedSecurityInfoW" (ByVal pObjectName As Long, ByVal ObjectType As Long, ByVal SecurityInfo As Long, ByVal psidOwner As Long, ByVal psidGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long) As Long
Private Declare Function GetAclInformation Lib "advapi32.dll" (ByVal pAcl As Long, ByVal pAclInformation As Long, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As ACL_INFORMATION_CLASS) As Long
Private Declare Function GetAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceIndex As Long, pAce As Long) As Long
Private Declare Function memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Private Declare Function GetExplicitEntriesFromAcl Lib "advapi32.dll" Alias "GetExplicitEntriesFromAclW" (ByVal pAcl As Long, pcCountOfExplicitEntries As Long, pListOfExplicitEntries As Long) As Long
Private Declare Function DeleteAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceIndex As Long) As Long
Private Declare Function InitializeAcl Lib "advapi32.dll" (ByVal pAcl As Long, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Private Declare Function LocalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function IsValidAcl Lib "advapi32.dll" (ByVal pAcl As Long) As Long
'Private Declare Function TreeResetNamedSecurityInfo Lib "advapi32.dll" Alias "TreeResetNamedSecurityInfoW" (ByVal pObjectName As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As SECURITY_INFORMATION, ByVal pOwner As Long, ByVal pGroup As Long, ByVal pDacl As Long, ByVal pSacl As Long, ByVal KeepExplicit As Long, ByVal fnProgress As Long, ByVal ProgressInvokeSetting As Long, ByVal Args As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Const MAX_KEYNAME            As Long = 255&

Private Const REG_OPTION_BACKUP_RESTORE As Long = 4&
Private Const GENERIC_ALL            As Long = &H10000000
Private Const GENERIC_READ           As Long = &H80000000
Private Const WRITE_DAC              As Long = &H40000
Private Const WRITE_OWNER            As Long = &H80000
Private Const READ_CONTROL           As Long = &H20000
Private Const KEY_WOW64_64KEY        As Long = &H100&
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY             As Long = 8&
Private Const SE_PRIVILEGE_ENABLED    As Long = 2&

Private Const OBJECT_INHERIT_ACE     As Long = 1&
Private Const CONTAINER_INHERIT_ACE  As Long = 2&

Private Const NO_MULTIPLE_TRUSTEE    As Long = 0&

Private Const ACCESS_DENIED_ACE_TYPE As Long = 1&

Private Const REG_CREATED_NEW_KEY    As Long = 1&

Private Const ERROR_MORE_DATA        As Long = 234&
Private Const ERROR_SUCCESS          As Long = 0&
Private Const ERROR_NO_TOKEN         As Long = 1008&

Private Const LMEM_FIXED             As Long = 0&
Private Const LMEM_ZEROINIT          As Long = &H40&

Private Const ACL_REVISION           As Long = 2&

Private Const ProgressInvokeNever    As Long = 1&

Private Const HKEY_CLASSES_ROOT      As Long = &H80000000
Private Const HKEY_CURRENT_USER      As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE     As Long = &H80000002
Private Const HKEY_USERS             As Long = &H80000003
Private Const HKEY_PERFORMANCE_DATA  As Long = &H80000004
Private Const HKEY_CURRENT_CONFIG    As Long = &H80000005
Private Const HKEY_DYN_DATA          As Long = &H80000006


' Creates array of EXPLICIT_ACCESS structures with access rights needed
Function Make_Default_Ace_Explicit(lHive As Long, KeyName As String) As EXPLICIT_ACCESS()
    
    Dim i As Long
    Dim cntAces As Long
    Dim idx As Long
    Dim pSid As Long
    Dim inf(68) As Long: inf(0) = 276: GetVersionEx inf(0)
    Dim MajorMinor As Single: MajorMinor = inf(1) + inf(2) / 10
    
    Static IsInit As Boolean
    
    Static bufSidAdmin()        As Byte
    Static bufSidSystem()       As Byte
    Static bufSidUsers()        As Byte
    Static bufSidPowerUsers()   As Byte
    Static bufSidRestricted()   As Byte
    Static bufSidCreator()      As Byte
    Static bufSidTI()           As Byte
    Static bufSidAppX()         As Byte
    
    If Not IsInit Then
        IsInit = True
        bufSidSystem = CreateBufferedSID("S-1-5-18")
        bufSidAdmin = CreateBufferedSID("S-1-5-32-544")
        bufSidUsers = CreateBufferedSID("S-1-5-32-545")
        bufSidPowerUsers = CreateBufferedSID("S-1-5-32-547") '( < Vista)
        bufSidRestricted = CreateBufferedSID("S-1-5-12")
        bufSidCreator = CreateBufferedSID("S-1-3-0")
        bufSidTI = CreateBufferedSID("S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464")  '(Win Vista+)
        bufSidAppX = CreateBufferedSID("S-1-15-2-1")
        
        'ÖÅÍÒÐ ÏÀÊÅÒÎÂ ÏÐÈËÎÆÅÍÈÉ\ÂÑÅ ÏÀÊÅÒÛ ÏÐÈËÎÆÅÍÈÉ (AppX) - S-1-15-2-1 (Win 8.0+)
        
        'TrustedInstaller - details on:
        '(EN) https://technet.microsoft.com/en-us/magazine/2007.06.acl.aspx
        '(RU) http://www.oszone.net/5003
        
        ' Well-known SIDs: https://support.microsoft.com/en-us/kb/243330
        '
        ' Other useful SIDs:
        '
        'Creator - S-1-3-0
        'Everyone = S-1-1-0
        'All Services = S-1-5-80-0
        'Local System - S-1-5-18
        'Local Administrator - S-1-5-21-500
        'Administrators - S-1-5-32-544
        'Users - S-1-5-32-545
        'Power Users = S-1-5-32-547
        'Guests = S-1-5-32-546
        'Restricted Code - S-1-5-12
        'Low Mandatory Level - S-1-16-4096
        'Medium Mandatory Level - S-1-16-8192
        'Medium Plus Mandatory Level - S-1-16-8448
        'High Mandatory Level - S-1-16-12288
        'System Mandatory Level - S-1-16-16384
        'Protected Process Mandatory Level - S-1-16-20480
        'Secure Process Mandatory Level - S-1-16-28672
        
    End If
    
    'array should be consistent
    ReDim Ace_Explicit(10) As EXPLICIT_ACCESS   '// now used 5-8/10
    
    '1. Local System:F (OI)(CI)
    idx = 0
    pSid = VarPtr(bufSidSystem(0))
    If IsValidSid(pSid) Then
      With Ace_Explicit(idx)
        .grfAccessPermissions = GENERIC_ALL
        .grfAccessMode = SET_ACCESS
        .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
        With .tTrustee
            .TrusteeForm = TRUSTEE_IS_SID
            .TrusteeType = TRUSTEE_IS_WELL_KNOWN_GROUP
            .ptstrName = pSid
        End With
      End With
      idx = idx + 1
    End If
    
    '2. Administrators:F (OI)(CI)
    pSid = VarPtr(bufSidAdmin(0))
    If IsValidSid(pSid) Then
      With Ace_Explicit(idx)
        .grfAccessPermissions = GENERIC_ALL
        .grfAccessMode = SET_ACCESS
        .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
        With .tTrustee
            .TrusteeForm = TRUSTEE_IS_SID
            .TrusteeType = TRUSTEE_IS_WELL_KNOWN_GROUP
            .ptstrName = pSid
        End With
      End With
      idx = idx + 1
    End If
    
    '3. Service:F (OI)(CI) (optional), depends on Key name
    If lHive = HKEY_LOCAL_MACHINE And InStr(1, KeyName, "SYSTEM\CurrentControlSet\services\", 1) = 1 Then
    
      Dim Tok, SrvName As String
      Tok = Split(KeyName, "\")
      If UBound(Tok) >= 3 Then SrvName = Tok(3)
    
      With Ace_Explicit(idx)
        .grfAccessPermissions = GENERIC_ALL
        .grfAccessMode = SET_ACCESS
        .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
        With .tTrustee
            .TrusteeForm = TRUSTEE_IS_NAME
            .TrusteeType = TRUSTEE_IS_UNKNOWN
            .ptstrName = StrPtr("NT SERVICE\" & SrvName)
        End With
      End With
      idx = idx + 1
    End If
    
    '4. Trusted Installer:F (OI)(CI) (optional) (Vista+)
    If MajorMinor >= 6 Then
      pSid = VarPtr(bufSidTI(0))
      If IsValidSid(pSid) Then
        With Ace_Explicit(idx)
          .grfAccessPermissions = GENERIC_ALL
          .grfAccessMode = SET_ACCESS
          .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
          With .tTrustee
            .TrusteeForm = TRUSTEE_IS_SID
            .TrusteeType = TRUSTEE_IS_UNKNOWN
            .ptstrName = pSid
          End With
        End With
        idx = idx + 1
      End If
    End If
    
    '5. AppX:R (OI)(CI) (optional) (Win 8.0+)
    If MajorMinor >= 6.2 Then
      pSid = VarPtr(bufSidAppX(0))
      If IsValidSid(pSid) Then
        With Ace_Explicit(idx)
          .grfAccessPermissions = GENERIC_READ
          .grfAccessMode = SET_ACCESS
          .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
          With .tTrustee
            .TrusteeForm = TRUSTEE_IS_SID
            .TrusteeType = TRUSTEE_IS_UNKNOWN
            .ptstrName = pSid
          End With
        End With
        idx = idx + 1
      End If
    End If
    
    '+ 2-3 "ACE" descriptions. Rights based on hive name - LM / CU.
    
    If lHive = HKEY_CURRENT_USER Then
      'HKCU
      'Users:F (OI)(CI)
      'Restricted:R (OI)(CI)
      pSid = VarPtr(bufSidUsers(0))
      If IsValidSid(pSid) Then
        With Ace_Explicit(idx)
          .grfAccessPermissions = GENERIC_ALL
          .grfAccessMode = SET_ACCESS
          .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
          With .tTrustee
            .TrusteeForm = TRUSTEE_IS_SID
            .TrusteeType = TRUSTEE_IS_WELL_KNOWN_GROUP
            .ptstrName = pSid
          End With
        End With
        idx = idx + 1
      End If
      
      pSid = VarPtr(bufSidRestricted(0))
      If IsValidSid(pSid) Then
        With Ace_Explicit(idx)
          .grfAccessPermissions = GENERIC_READ
          .grfAccessMode = SET_ACCESS
          .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
          With .tTrustee
            .TrusteeForm = TRUSTEE_IS_SID
            .TrusteeType = TRUSTEE_IS_WELL_KNOWN_GROUP
            .ptstrName = pSid
          End With
        End With
        idx = idx + 1
      End If
      
    Else
      'HKLM
      'Creator:F (CI)
      'Users:R (OI)(CI)
      'PowerUsers:R (OI)(CI) (XP only)
      pSid = VarPtr(bufSidCreator(0))
      If IsValidSid(pSid) Then
        With Ace_Explicit(idx)
          .grfAccessPermissions = GENERIC_ALL
          .grfAccessMode = SET_ACCESS
          .grfInheritance = CONTAINER_INHERIT_ACE
          With .tTrustee
            .TrusteeForm = TRUSTEE_IS_SID
            .TrusteeType = TRUSTEE_IS_WELL_KNOWN_GROUP
            .ptstrName = pSid
          End With
        End With
        idx = idx + 1
      End If
      
      pSid = VarPtr(bufSidUsers(0))
      If IsValidSid(pSid) Then
        With Ace_Explicit(idx)
          .grfAccessPermissions = GENERIC_READ
          .grfAccessMode = SET_ACCESS
          .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
          With .tTrustee
            .TrusteeForm = TRUSTEE_IS_SID
            .TrusteeType = TRUSTEE_IS_WELL_KNOWN_GROUP
            .ptstrName = pSid
          End With
        End With
        idx = idx + 1
      End If
      
      If MajorMinor < 6 Then
        pSid = VarPtr(bufSidPowerUsers(0))
        If IsValidSid(pSid) Then
          With Ace_Explicit(idx)
            .grfAccessPermissions = GENERIC_READ
            .grfAccessMode = SET_ACCESS
            .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
            With .tTrustee
              .TrusteeForm = TRUSTEE_IS_SID
              .TrusteeType = TRUSTEE_IS_WELL_KNOWN_GROUP
              .ptstrName = pSid
            End With
          End With
          idx = idx + 1
        End If
      End If
      
    End If
    
    If idx > 0 Then
        ReDim Preserve Ace_Explicit(idx - 1)
    End If
    
    Make_Default_Ace_Explicit = Ace_Explicit
End Function


'Creates SID buffer array from SID string
Public Function CreateBufferedSID(SidString As String) As Byte()
    Dim pSid        As Long
    Dim cbSID       As Long

    ReDim bufSid(0) As Byte

    If 0 = ConvertStringSidToSid(StrPtr(SidString), pSid) Then  ' * -> *
        Debug.Print "ErrorHandler: ConvertStringSidToSidW failed. Input buffer: " & SidString
    Else
        If IsValidSid(pSid) Then
            cbSID = GetLengthSid(pSid)
    
            If cbSID <> 0 Then
                ReDim bufSid(cbSID - 1) As Byte
                CopySid cbSID, VarPtr(bufSid(0)), pSid
            End If
    
            LocalFree pSid
        End If
        CreateBufferedSID = bufSid
    End If

End Function

' if main hive handle wasn't defined, assigns handle according to hive's name defined by Full key name directed
Sub NormalizeKeyNameAndHiveHandle(ByRef lHive As Long, ByRef KeyName As String)
    Dim iPos        As Long
    If lHive = 0 Then
        lHive = GetHKey(KeyName)
        iPos = InStr(KeyName, "\")
        If (iPos <> 0&) Then KeyName = Mid$(KeyName, iPos + 1&) Else KeyName = vbNullString
    End If
End Sub

''concat structure 'EXPLICIT_ACCESS' to array in consistent order
'Function Add_Ace_Explicit(ByRef Ace_Explicit() As EXPLICIT_ACCESS, New_Ace_Explicit As EXPLICIT_ACCESS)
'    Dim i As Long
'    ReDim Concat_Ace_Explicit(UBound(Ace_Explicit) + 1) As EXPLICIT_ACCESS
'
'    For i = 0 To UBound(Ace_Explicit)   'duplicate old array
'        Concat_Ace_Explicit(i) = Ace_Explicit(i)
'    Next
'    Concat_Ace_Explicit(UBound(Ace_Explicit) + 1) = New_Ace_Explicit
'    'Replacing array
'    Ace_Explicit = Concat_Ace_Explicit
'End Function


' obtains ownership on registry key

Public Function RegKeySetOwnerShip(lHive&, ByVal KeyName$, SidString As String, Optional bUseWow64 As Boolean) As Boolean
    '
    'Parameters:
    '
    'lHive - pseudohandle to root key (hive). This value can be 0.
    'KeyName - Path to registry key. Is lHive is 0, this path must be Full, otherwise it must be relative to hive.
    'SidString - string representation of SID
    'bUseWow64 - (optional) if true, this function use registry redirector, so all calls will be directed to 32-bit keys on 64-bit machine

    On Error GoTo ErrorHandler:
    
    Dim flagDisposition As Long
    Dim bufSid()    As Byte
    Dim hKey        As Long
    Dim lret        As Long
    
    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    ' -->>> moved to main form
'
'    SetCurrentProcessPrivileges "SeBackupPrivilege"
'    SetCurrentProcessPrivileges "SeRestorePrivilege"
'    SetCurrentProcessPrivileges "SeTakeOwnershipPrivilege"
    
    ' SeTakeOwnershipPrivilege + WRITE_OWNER
    If ERROR_SUCCESS <> RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, WRITE_OWNER Or (KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        'Key doesn't exist
        Exit Function
    Else
        RegCloseKey hKey
    End If
    
    bufSid = CreateBufferedSID(SidString)
    
    ' ACCESS_SYSTEM_SECURITY
    If ERROR_SUCCESS = RegCreateKeyEx(lHive, StrPtr(KeyName), 0&, 0&, REG_OPTION_BACKUP_RESTORE, WRITE_DAC Or (KEY_WOW64_64KEY And Not bUseWow64), ByVal 0&, hKey, flagDisposition) Then
        
'        If flagDisposition = REG_CREATED_NEW_KEY Then
'            RegCloseKey hKey
'            RegDeleteKeyEx lHive, StrPtr(KeyName), KEY_WOW64_64KEY And Not bUseWow64, 0&
'            Debug.Print "Key doesn't exist"
'            Exit Function
'        End If
        
        'IIf(bUseWow64 And isWin64(), SE_REGISTRY_WOW64_32KEY, SE_REGISTRY_KEY)
        
        lret = SetSecurityInfo(hKey, SE_REGISTRY_KEY, OWNER_SECURITY_INFORMATION, VarPtr(bufSid(0)), 0&, 0&, 0&)
        
        If lret = ERROR_SUCCESS Then
            
            RegKeySetOwnerShip = True
            Debug.Print KeyName & " - OwnerShip granted successfully."
        
        Else

            Debug.Print KeyName & " - Error in SetSecurityInfo: " & lret
            
        End If
        
        RegCloseKey hKey
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print "Error in RegSetOwnerShip", err, err.Description
End Function


'resets access on registry key into some standart (look below for details)

Public Function RegKeyResetDACL(lHive&, ByVal KeyName$, Optional bUseWow64 As Boolean, Optional Recursive As Boolean = False) As Boolean
    '
    'Parameters:
    '
    'lHive - pseudohandle to root key (hive). This value can be 0.
    'KeyName - Path to registry key. Is lHive is 0, this path must be Full, otherwise it must be relative to hive.
    'bUseWow64 - (optional) if true, this function use registry redirector, so all calls will be directed to 32-bit keys on 64-bit machine
    'Recursive - (optional) apply action to all subkeys.

    On Error GoTo ErrorHandler:
    'Description:
    '
    'This function also made call to RegKeySetOwnerShip function.
    'So, you don't need to call it twice.
    '
    'Note:
    'There are 6 types of DACL ACEs: 3 of them - General, other 3 - Object specific (has more fields on its struct, incl. GUIDs)
    'This function performs check of ACCESS_DENIED_ACE presence.
    '
    'If DACL contains such struct, this ACE will be removed from it.
    'So, all denied access rights will be revoked.
    '
    'Default access rights will be written using EXPLICIT_ACCESS array supplemented by default access masks for standart trustees
    'like 'Local System', 'Administrators' and so on, see details on top: Make_Default_Ace_Explicit function.
    '
    'EXPLICIT_ACCESS will be applied by merging to ACL in consistent order using SetEntriesInAcl function.
    '
    
    Dim flagDisposition As Long
    Dim hKey        As Long
    Dim RelSD()     As Byte
    Dim AbsSD()     As Byte
    Dim cbRelSD     As Long
    Dim cbAbsSD     As Long
    Dim oldDACL()   As Byte
    Dim newDACL()   As Byte
    Dim cbDACL      As Long
    Dim cbSACL      As Long
    Dim cbSID_Owner As Long
    Dim cbPrimGrp   As Long
    Dim pNewDacl    As Long
    Dim AclInfo     As ACL_SIZE_INFORMATION
    Dim AceDenied   As ACCESS_DENIED_ACE
    Dim AceHead     As ACE_HEADER
    Dim i           As Long
    Dim pAce        As Long
    Dim lret        As Long
    Dim Revoke_Ace_Explicit As EXPLICIT_ACCESS
    Dim cExplicit   As Long
    Dim pListExplicit As Long
    Dim pAcl        As Long
    Dim hKeyEnum    As Long
    Dim sSubKeyName As String
    
    Call NormalizeKeyNameAndHiveHandle(lHive, KeyName)
    
    RegKeySetOwnerShip lHive, KeyName, "S-1-5-32-544", bUseWow64
    
    ' -->>> moved to main form
'
'    SetCurrentProcessPrivileges "SeBackupPrivilege"
'    SetCurrentProcessPrivileges "SeRestorePrivilege"
'    SetCurrentProcessPrivileges "SeTakeOwnershipPrivilege"
'    SetCurrentProcessPrivileges "SeSecurityPrivilege"       'SACL
    
    If ERROR_SUCCESS <> RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, WRITE_OWNER Or (KEY_WOW64_64KEY And Not bUseWow64), hKey) Then
        'Key doesn't exist
        Exit Function
    Else
        RegCloseKey hKey
    End If
    
    If ERROR_SUCCESS = RegCreateKeyEx(lHive, StrPtr(KeyName), 0&, 0&, _
        REG_OPTION_BACKUP_RESTORE, _
        READ_CONTROL Or WRITE_DAC Or (KEY_WOW64_64KEY And Not bUseWow64), _
        ByVal 0&, hKey, flagDisposition) Then
    
        ReDim RelSD(0)
        
        'extracting relative SD
        
        GetKernelObjectSecurity hKey, DACL_SECURITY_INFORMATION, VarPtr(RelSD(0)), 0&, VarPtr(cbRelSD)
        
        If cbRelSD <> 0 Then
        
            ReDim RelSD(cbRelSD - 1)
            
            If GetKernelObjectSecurity(hKey, DACL_SECURITY_INFORMATION, VarPtr(RelSD(0)), cbRelSD, VarPtr(cbRelSD)) Then
        
                'relative SD -> Absolute
        
                MakeAbsoluteSD VarPtr(RelSD(0)), 0&, VarPtr(cbAbsSD), 0&, VarPtr(cbDACL), 0&, VarPtr(cbSACL), 0&, VarPtr(cbSID_Owner), 0&, VarPtr(cbPrimGrp)
                
                ReDim AbsSD(cbAbsSD - 1)
                If cbDACL <> 0 Then
                    ReDim oldDACL(cbDACL - 1)
                Else 'if not contains DACL information
                    ReDim oldDACL(0)
                End If
                
                If MakeAbsoluteSD(VarPtr(RelSD(0)), VarPtr(AbsSD(0)), VarPtr(cbAbsSD), VarPtr(oldDACL(0)), VarPtr(cbDACL), 0&, VarPtr(cbSACL), 0&, VarPtr(cbSID_Owner), 0&, VarPtr(cbPrimGrp)) Then
                
                    If IsValidSecurityDescriptor(VarPtr(AbsSD(0))) Then
                    
                        'making default ACE descriptions
                        
                        Dim Ace_Explicit() As EXPLICIT_ACCESS
                        
                        Ace_Explicit = Make_Default_Ace_Explicit(lHive, KeyName)
                        
'                        'appending it with revoking ACE descriptions of those SIDs who currently has denied access rights on SD
'
'                        'LookupSecurityDescriptorParts (if need SACL / DACL)
'
'                        If ERROR_SUCCESS = GetExplicitEntriesFromAcl(VarPtr(oldDACL(0)), cExplicit, pListExplicit) Then
'
'                            For i = 0 To cExplicit - 1
'
'                                memcpy Revoke_Ace_Explicit, ByVal (pListExplicit + LenB(Revoke_Ace_Explicit) * i), LenB(Revoke_Ace_Explicit)
'
'                                If Revoke_Ace_Explicit.grfAccessMode = DENY_ACCESS Then
'
'                                    Revoke_Ace_Explicit.grfAccessMode = GRANT_ACCESS   ' REVOKE_ACCESS
'
'                                    ' rebuild array into consistent order
'                                    Call Add_Ace_Explicit(Ace_Explicit(), Revoke_Ace_Explicit)
'
'                                End If
'                            Next
'
'                            LocalFree pListExplicit
'                        End If
                        
                        'appending it with revoking ACE descriptions of those SIDs who currently has denied access rights on SD
                        
                        If cbDACL > 0 Then
                          'has DACL
                          If GetAclInformation(VarPtr(oldDACL(0)), VarPtr(AclInfo), LenB(AclInfo), AclSizeInformation) Then

                            For i = AclInfo.AceCount - 1 To 0 Step -1

                                If GetAce(VarPtr(oldDACL(0)), i, pAce) Then

                                    memcpy AceHead, ByVal pAce, LenB(AceHead)   ' ((ACE_HEADER) pAce) -> AceType

                                    If AceHead.AceType = ACCESS_DENIED_ACE_TYPE Then

                                        lret = DeleteAce(VarPtr(oldDACL(0)), i)

                                        'old routine - revoking access (but SetEntriesInAcl doesn't support it for ACCESS_DENIED_ACE type)
'                                        'memcpy AceDenied, ByVal pAce, LenB(AceDenied)
'                                        'Debug.Print AceDenied.SidStart
'
'                                        'SidStart contains first DWORD of SID buffer.
'                                        'Rest part is stored directly behind the structure.
'                                        'So, ptr to SidStart can be used (its offset = 0x8)
'
'                                        With Revoke_Ace_Explicit
'
'                                            .grfAccessPermissions = GENERIC_ALL
'                                            .grfAccessMode = REVOKE_ACCESS ' SET_ACCESS
'                                            .grfInheritance = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
'                                            With .tTrustee
'                                                .TrusteeForm = TRUSTEE_IS_SID
'                                                .TrusteeType = TRUSTEE_IS_UNKNOWN
'                                                .ptstrName = pAce + 8&
'                                            End With
'                                        End With
'
'                                        ' rebuild array into consistent order
'                                        Call Add_Ace_Explicit(Ace_Explicit(), Revoke_Ace_Explicit)
                                    End If

                                End If
                            Next
                          End If
                        End If
                        
                        lret = -1
                        
                        If cbDACL = 0 Then
                            pAcl = CreateEmptyACL(Ace_Explicit)
                        Else
                            pAcl = VarPtr(oldDACL(0))
                        End If
                        
                        'merging ACE descriptions with existed DACL
                        If IsValidAcl(pAcl) Then
                            lret = SetEntriesInAcl(UBound(Ace_Explicit) + 1, VarPtr(Ace_Explicit(0)), pAcl, pNewDacl)
                        End If
                        
                        If cbDACL = 0 Then LocalFree pAcl
                        
                        If cbDACL > 0 And ERROR_SUCCESS <> lret Then
                            'for instance, not enough quota -> making DACL from default ACE_EXPLICIT
                            
                            pAcl = CreateEmptyACL(Ace_Explicit)
                            
                            If IsValidAcl(pAcl) Then
                                lret = SetEntriesInAcl(UBound(Ace_Explicit) + 1, VarPtr(Ace_Explicit(0)), pAcl, pNewDacl)
                            End If
                        
                            LocalFree pAcl
                        End If
                        
                        If ERROR_SUCCESS = lret Then
                            
                            'apply it
                    
'                            If ERROR_SUCCESS = SetNamedSecurityInfo( _
'                                StrPtr(ConvertHiveHandleToSeObjectName(lHive) & "\" & KeyName), _
'                                SE_REGISTRY_KEY, DACL_SECURITY_INFORMATION, 0&, 0&, pNewDacl, 0&) Then
'
'                                Debug.Print "Permissions granted successfully."
'                            End If

                            'x64 support
                            '+ protected DACL (prevent DACL to inherite ACEs from parent)
                            
                            'IIf(bUseWow64 And isWin64(), SE_REGISTRY_WOW64_32KEY, SE_REGISTRY_KEY)
                            
                            If ERROR_SUCCESS = SetSecurityInfo(hKey, SE_REGISTRY_KEY, _
                                DACL_SECURITY_INFORMATION Or PROTECTED_DACL_SECURITY_INFORMATION, _
                                0&, 0&, pNewDacl, 0&) Then
                                
                                RegKeyResetDACL = True
                                Debug.Print KeyName & " - Permissions granted successfully."
                                
                                If Recursive Then
                                
                                    'This 'tree' function produces some strange ACEs with duplicate records of inheritance from parent objects.
                                    'Besides, inherintance affects by grand parent objects too that may be not a part of fix.
                                    'For instance, it may cause apllying denied access from much higher group like well-known SID 'All'.
                                    'As a result, access to any objects still be denied.
                                    '
                                    'So, TreeResetNamedSecurityInfo is not an option here, I guess.
                                    'And explicit ACE with manual enumaration of all subkeys must be much better decision.
                                    'It will be compatible with 64-keys too.
                                
'                                    'WOW64_64 not supported
'                                    lret = TreeResetNamedSecurityInfo( _
'                                        StrPtr(ConvertHiveHandleToSeObjectName(lHive) & "\" & KeyName), _
'                                        SE_REGISTRY_KEY, DACL_SECURITY_INFORMATION, 0&, 0&, pNewDacl, 0&, CLng(False), 0&, ProgressInvokeNever, 0&)
'
'                                    If lret = ERROR_SUCCESS Then
'                                        Debug.Print "Permissions on tree granted successfully."
'                                        RegKeyResetDACL = True
'                                    Else
'                                        RegKeyResetDACL = False
'                                    End If

                                    'Ùå íå âìåðëà Óêðà¿íà :)
                                    If RegOpenKeyEx(lHive, StrPtr(KeyName), 0&, KEY_ENUMERATE_SUB_KEYS Or (KEY_WOW64_64KEY And Not bUseWow64), hKeyEnum) = ERROR_SUCCESS Then
    
                                        sSubKeyName = String$(MAX_KEYNAME, vbNullChar)
                                        
                                        i = 0
                                        Do While RegEnumKeyEx(hKeyEnum, i, StrPtr(sSubKeyName), MAX_KEYNAME, 0&, 0&, 0&, ByVal 0&) = ERROR_SUCCESS
                                        
                                            sSubKeyName = Left$(sSubKeyName, lstrlen(StrPtr(sSubKeyName)))
                                            
                                            RegKeyResetDACL lHive, KeyName & IIf(0 <> Len(KeyName), "\", "") & sSubKeyName, bUseWow64, True
                                            
                                            sSubKeyName = String$(MAX_KEYNAME, vbNullChar)
                                            i = i + 1
                                        Loop
                                        RegCloseKey hKeyEnum
                                    End If
            
                                End If

                            End If
                            
                            LocalFree pNewDacl
                    
                        End If
                        
                    End If
                
                End If
        
            End If
        End If
    
        RegCloseKey hKey
    
    End If
    
    If Not RegKeyResetDACL Then Debug.Print KeyName & " - Failed to grant permissions!"

    Exit Function
ErrorHandler:
    Debug.Print "Error in SetDACL", err, err.Description
End Function

'returns ptr to new ACL
'size of ACL calculated from array of EXPLICIT_ACCESS
Function CreateEmptyACL(Ace_Explicit() As EXPLICIT_ACCESS) As Long
    Dim pAcl As Long
    Dim cbACL As Long
    Dim Num_of_ACEs As Long
    Dim i As Long
    
    Num_of_ACEs = UBound(Ace_Explicit) + 1
    
    cbACL = 8& + (12& * Num_of_ACEs)     'sizeof(ACL) + (NUM_OF_ACES * sizeof(ACCESS_ALLOWED_ACE))
    
    For i = 0 To Num_of_ACEs - 1
        
        If IsValidSid(Ace_Explicit(i).tTrustee.ptstrName) Then
            
            cbACL = cbACL + GetLengthSid(Ace_Explicit(i).tTrustee.ptstrName) - 4&   ' - sizeof(DWORD)
            
        End If
    Next
    
    'Align cbAcl to a DWORD
    cbACL = (cbACL + 3) And &HFFFFFFFC  ' 3 = sizeof(DWORD) - 1)
    
    pAcl = LocalAlloc(LMEM_FIXED Or LMEM_ZEROINIT, cbACL)
    
    If pAcl <> 0 Then
                            
        If InitializeAcl(pAcl, cbACL, ACL_REVISION) Then
        
            CreateEmptyACL = pAcl
        
        End If
                            
    End If
End Function

Private Function GetHKey(ByVal HKeyName As String) As Long 'Get handle of main hive
    On Error GoTo ErrorHandler:
    Dim pos As Long
    pos = InStr(HKeyName, "\")
    If pos <> 0 Then HKeyName = Left$(HKeyName, pos - 1)
    Select Case UCase$(HKeyName)
        Case "HKEY_CLASSES_ROOT", "HKCR"
            GetHKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER", "HKCU"
            GetHKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE", "HKLM"
            GetHKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS", "HKU"
            GetHKey = HKEY_USERS
        Case "HKEY_PERFORMANCE_DATA"
            GetHKey = HKEY_PERFORMANCE_DATA
        Case "HKEY_CURRENT_CONFIG", "HKCC"
            GetHKey = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            GetHKey = HKEY_DYN_DATA
    End Select
    Exit Function
ErrorHandler:
    Debug.Print "Error in GetHKey", err, err.Description
    If inIDE Then Stop: Resume Next
End Function

Private Function ConvertHiveHandleToSeObjectName(hHive As Long) As String
    Dim SeObj As String
    Select Case hHive
    Case &H80000000
        SeObj = "CLASSES_ROOT"
    Case &H80000001
        SeObj = "CURRENT_USER"
    Case &H80000002
        SeObj = "MACHINE"
    Case &H80000003
        SeObj = "USERS"
    End Select
    ConvertHiveHandleToSeObjectName = SeObj
End Function

Public Function SetCurrentProcessPrivileges(PrivilegeName As String) As Boolean
    
    Dim tp As TOKEN_PRIVILEGES, hToken&
    
    If LookupPrivilegeValue(0&, StrPtr(PrivilegeName), tp.LuidLowPart) Then   'i.e. "SeDebugPrivilege"
    
        If 0 = OpenThreadToken(GetCurrentThread(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, 1&, hToken) Then
        
            If err.LastDllError = ERROR_NO_TOKEN Then
            
                If 0 = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If

        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
        SetCurrentProcessPrivileges = AdjustTokenPrivileges(hToken, 0&, tp, 0&, 0&, 0&)
        CloseHandle hToken
    End If
End Function

Function IsWin64() As Boolean
    Const PROCESSOR_ARCHITECTURE_AMD64 As Long = 9&
    Dim si(35) As Byte
    GetNativeSystemInfo VarPtr(si(0))
    If si(0) And PROCESSOR_ARCHITECTURE_AMD64 Then IsWin64 = True
End Function
