Attribute VB_Name = "modAccount"
Option Explicit

'see:
'https://learn.microsoft.com/en-us/windows/win32/secauthz/sid-components
'https://learn.microsoft.com/en-us/windows/win32/secauthz/well-known-sids
'https://github.com/winsiderss/systeminformer/blob/master/phlib/include/phnative.h#L870

Public Enum LSA_USER_ACCOUNT_TYPE
    UnknownUserAccountType
    LocalUserAccountType
    PrimaryDomainUserAccountType
    ExternalDomainUserAccountType
    LocalConnectedUserAccountType
    AADUserAccountType
    InternetUserAccountType
    MSAUserAccountType
End Enum

Public Declare Function LsaLookupUserAccountType Lib "sechost.dll" (ByVal pSid As Long, accountType As LSA_USER_ACCOUNT_TYPE) As Long

Private Function Wrap_LsaLookupUserAccountType(pSid As Long, out_AccountType As LSA_USER_ACCOUNT_TYPE) As Long
    On Error GoTo ErrorHandler:
    
    Static bInit As Boolean
    Dim status As Long
    Dim accountType As LSA_USER_ACCOUNT_TYPE
    
    If Not IsProcedureAvail("LsaLookupUserAccountType", "sechost.dll") Then
        Wrap_LsaLookupUserAccountType = STATUS_UNSUCCESSFUL
        Exit Function
    End If
    
    status = LsaLookupUserAccountType(pSid, accountType)

    If (NT_SUCCESS(status)) Then
        out_AccountType = accountType
    End If

    Wrap_LsaLookupUserAccountType = status
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Wrap_LsaLookupUserAccountType"
    If inIDE Then Stop: Resume Next
End Function

'thanks to dmex
Private Function GetUserAccountTypeBySidPtr(pSid As Long) As String
    
    Dim accountType As LSA_USER_ACCOUNT_TYPE
    Dim accountTypeStr As String
    
    If (NT_SUCCESS(Wrap_LsaLookupUserAccountType(pSid, accountType))) Then
        Select Case (accountType)
        Case LocalUserAccountType: accountTypeStr = "Local"
        Case PrimaryDomainUserAccountType: accountTypeStr = "ActiveDirectory"
        Case ExternalDomainUserAccountType: accountTypeStr = "ActiveDirectory"
        Case LocalConnectedUserAccountType: accountTypeStr = "Microsoft"
        Case MSAUserAccountType: accountTypeStr = "Microsoft"
        Case AADUserAccountType: accountTypeStr = "AzureAD"
        Case InternetUserAccountType:
            Dim pSidMsaAuthority As Long
            Dim tpSidAuth As SID_IDENTIFIER_AUTHORITY
            tpSidAuth.Value(5) = SECURITY_AUTHENTICATED_USER_RID
            
            Call AllocateAndInitializeSid(tpSidAuth, 2, 0, 0, 0, 0, 0, 0, 0, 0, pSidMsaAuthority)
            
            If EqualPrefixSid(pSid, pSidMsaAuthority) Then
                accountTypeStr = "Microsoft"
            Else
                accountTypeStr = "Internet"
            End If
            
            FreeSid pSidMsaAuthority
        End Select
    End If
    
    GetUserAccountTypeBySidPtr = accountTypeStr
    
End Function

Private Function GetUserAccountTypeBySidString(sSid As String) As String
    Dim bufSid() As Byte
    Dim pSid As Long
    bufSid = CreateBufferedSID(sSid)
    If AryPtr(bufSid) <> 0 Then
        pSid = VarPtr(bufSid(0))
        GetUserAccountTypeBySidString = GetUserAccountTypeBySidPtr(pSid)
    End If
End Function

Public Function GetCurrentUserAccountType() As String
    If OSver.IsWindows8OrGreater Then
        GetCurrentUserAccountType = GetUserAccountTypeBySidString(Reg.CurrentUserSID)
    End If
End Function

Public Function GetUserAccountType(sUser As String) As String
    If OSver.IsWindows8OrGreater Then
        Dim sSid As String
        sSid = GetUserSid(sUser)
        If Len(sSid) <> 0 Then GetUserAccountType = GetUserAccountTypeBySidString(sSid)
    End If
End Function

Public Function GetUserSid(sUser As String) As String
    Dim sThisUser As String
    Dim aSubKeys() As String
    Dim i As Long
    For i = 1 To Reg.EnumSubKeysToArray(HKEY_USERS, "", aSubKeys())
        If aSubKeys(i) Like "S-#-#-#*" And Not StrEndWith(aSubKeys(i), "_Classes") Then
            sThisUser = MapSIDToUsername(aSubKeys(i))
            If StrComp(sUser, sThisUser, vbTextCompare) = 0 Then
                GetUserSid = aSubKeys(i)
                Exit For
            End If
        End If
    Next i
End Function

Public Function GetLocalGroupNames() As String()
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetLocalGroupNames - Begin"
    
    Dim nStatus As Long
    Dim dwLevel As Long
    Dim groupinfo As LOCALGROUP_INFO_0
    Dim dwEntriesRead As Long
    Dim dwTotalEntries As Long
    Dim dwResumeHandle As Long
    Dim pBuf As Long, pTmpBuf As Long
    Dim i As Long
    
    dwLevel = 0 'for LOCALGROUP_INFO_0
    
    Do
        nStatus = NetLocalGroupEnum(0&, dwLevel, pBuf, MAX_PREFERRED_LENGTH, dwEntriesRead, dwTotalEntries, dwResumeHandle)
        
        If nStatus = NERR_Success Or nStatus = ERROR_MORE_DATA Then
            If pBuf <> 0 Then
                pTmpBuf = pBuf
                For i = 0 To dwEntriesRead - 1
                    'memcpy userinfo, ByVal pTmpBuf, LenB(userinfo)
                    GetMem4 ByVal pTmpBuf, groupinfo
                    ArrayAddStr GetLocalGroupNames, StringFromPtrW(groupinfo.lgrpi0_name)
                    pTmpBuf = pTmpBuf + LenB(groupinfo)
                Next
            End If
        End If
        
        If pBuf Then
            NetApiBufferFree pBuf
        End If
        
        If pBuf = 0 Or dwResumeHandle = 0 Then Exit Do
        
    Loop While nStatus = ERROR_MORE_DATA
    
    AppendErrorLogCustom "GetLocalGroupNames - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetLocalGroupNames"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetUserGroupNames(sUser As String) As String()
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetUserGroupNames - Begin"
    
    Dim nStatus As Long
    Dim dwLevel As Long
    Dim groupinfo As LOCALGROUP_INFO_0
    Dim dwEntriesRead As Long
    Dim dwTotalEntries As Long
    Dim dwResumeHandle As Long
    Dim pBuf As Long, pTmpBuf As Long
    Dim i As Long
    
    dwLevel = 0 'for LOCALGROUP_INFO_0
    
    nStatus = NetUserGetLocalGroups(0&, StrPtr(sUser), dwLevel, LG_INCLUDE_INDIRECT, pBuf, MAX_PREFERRED_LENGTH, dwEntriesRead, dwTotalEntries)
    
    If nStatus = NERR_Success Or nStatus = ERROR_MORE_DATA Then
        If pBuf <> 0 Then
            pTmpBuf = pBuf
            For i = 0 To dwEntriesRead - 1
                GetMem4 ByVal pTmpBuf, groupinfo
                ArrayAddStr GetUserGroupNames, StringFromPtrW(groupinfo.lgrpi0_name)
                pTmpBuf = pTmpBuf + LenB(groupinfo)
            Next
        End If
    End If
    
    If pBuf Then
        NetApiBufferFree pBuf
    End If
    
    AppendErrorLogCustom "GetUserGroupNames - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetUserGroupNames"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetLocalUserNames() As String()
    
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetLocalUserNames - Begin"
    
    Dim nStatus As Long
    Dim dwLevel As Long
    Dim userinfo As USER_INFO_0
    Dim dwEntriesRead As Long
    Dim dwTotalEntries As Long
    Dim dwResumeHandle As Long
    Dim pBuf As Long, pTmpBuf As Long
    Dim i As Long
    
    dwLevel = 0 'for USER_INFO_0
    
    Do
        nStatus = NetUserEnum(0&, dwLevel, FILTER_NORMAL_ACCOUNT, pBuf, MAX_PREFERRED_LENGTH, dwEntriesRead, dwTotalEntries, dwResumeHandle)
        
        If nStatus = NERR_Success Or nStatus = ERROR_MORE_DATA Then
            If pBuf <> 0 Then
                pTmpBuf = pBuf
                For i = 0 To dwEntriesRead - 1
                    'memcpy userinfo, ByVal pTmpBuf, LenB(userinfo)
                    GetMem4 ByVal pTmpBuf, userinfo
                    ArrayAddStr GetLocalUserNames, StringFromPtrW(userinfo.usri0_name)
                    pTmpBuf = pTmpBuf + LenB(userinfo)
                Next
            End If
        End If
        
        If pBuf Then
            NetApiBufferFree pBuf
        End If
        
        If pBuf = 0 Or dwResumeHandle = 0 Then Exit Do
        
    Loop While nStatus = ERROR_MORE_DATA

    AppendErrorLogCustom "GetLocalUserNames - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetLocalUserNames"
    If inIDE Then Stop: Resume Next
End Function

Public Function IsValidSidEx(sSid As String) As Boolean
    Dim bufSid() As Byte
    bufSid = CreateBufferedSID(sSid)
    IsValidSidEx = IsValidSid(VarPtr(bufSid(0)))
End Function

Public Function IsValidUserName(sUsername As String) As Boolean
    If Len(sUsername) = 0 Then Exit Function
    IsValidUserName = InArray(sUsername, g_LocalUserNames, , , vbTextCompare)
End Function

Public Function IsValidGroupName(sGroupname As String) As Boolean
    If Len(sGroupname) = 0 Then Exit Function
    IsValidGroupName = InArray(sGroupname, g_LocalGroupNames, , , vbTextCompare)
End Function

Public Function IsUserMembershipRDP(sUsername As String) As Boolean
    Dim sRdpSid As String, sRdpGroup As String
    sRdpSid = "S-1-5-32-555"
    sRdpGroup = MapSIDToUsername(sRdpSid)
    If Len(sRdpGroup) <> 0 Then
        IsUserMembershipRDP = IsUserMembershipInGroup(sUsername, sRdpGroup)
    End If
End Function

Public Function IsUserMembershipInGroup(sUsername As String, sGroup As String) As Boolean
    Dim aGroups() As String
    Dim i As Long
    aGroups = modAccount.GetUserGroupNames(sUsername)
    For i = 0 To UBoundSafe(aGroups)
        If StrComp(aGroups(i), sGroup, vbTextCompare) = 0 Then
            IsUserMembershipInGroup = True
            Exit For
        End If
    Next
End Function

Public Function RemoveUserGroupMembership(sUsername As String, sGroup As String) As Boolean
    RemoveUserGroupMembership = (NERR_Success = NetLocalGroupDelMembers(0&, StrPtr(sGroup), 3&, VarPtr(StrPtr(sUsername)), 1&)) 'LOCALGROUP_MEMBERS_INFO_3
End Function

Public Function AddUserGroupMembership(sUsername As String, sGroup As String) As Boolean
    AddUserGroupMembership = (NERR_Success = NetLocalGroupAddMembers(0&, StrPtr(sGroup), 3&, VarPtr(StrPtr(sUsername)), 1&)) 'LOCALGROUP_MEMBERS_INFO_3
End Function
