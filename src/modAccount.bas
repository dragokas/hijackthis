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
