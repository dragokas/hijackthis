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

Public Function GetSidAccountType(pSid As Long, out_AccountType As LSA_USER_ACCOUNT_TYPE) As Long
    On Error GoTo ErrorHandler:
    
    Dim status As Long
    Dim accountType As LSA_USER_ACCOUNT_TYPE
    
    If Not IsProcedureAvail("LsaLookupUserAccountType", "sechost.dll") Then
        GetSidAccountType = STATUS_UNSUCCESSFUL
        Exit Function
    End If
    
    status = LsaLookupUserAccountType(pSid, accountType)

    If (NT_SUCCESS(status)) Then
        out_AccountType = accountType
    End If

    GetSidAccountType = status
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetSidAccountType"
    If inIDE Then Stop: Resume Next
End Function


'thanks to dmex
Public Function GetSidAccountTypeString() As String
    
    Dim accountType As LSA_USER_ACCOUNT_TYPE
    Dim accountTypeStr As String
    
    Dim bufSid() As Byte
    Dim pSid As Long
    Dim sSid As String
    sSid = Reg.CurrentUserSID
    If Len(sSid) = 0 Then Exit Function
    bufSid = CreateBufferedSID(sSid)
    If AryPtr(bufSid) = 0 Then Exit Function
    pSid = VarPtr(bufSid(0))
    
    If (NT_SUCCESS(GetSidAccountType(pSid, accountType))) Then
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
    
    GetSidAccountTypeString = accountTypeStr
    
End Function
