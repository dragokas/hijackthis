Attribute VB_Name = "modOSInfo"
Option Explicit

Private Type RTL_OSVERSIONINFOEXW
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(127) As Integer
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

Private Type SID_AND_ATTRIBUTES
    SID As Long
    Attributes As Long
End Type

Private Type TOKEN_GROUPS
    GroupCount As Long
    Groups(20) As SID_AND_ATTRIBUTES
End Type

'Private Type TOKEN_PRIVILEGES
'    PrivilegeCount  As Long
'    LuidLowPart     As Long
'    LuidHighPart    As Long
'    Attributes      As Long
'End Type


Private Declare Function RtlGetVersion Lib "NTDLL.DLL" (lpVersionInformation As RTL_OSVERSIONINFOEXW) As Long
Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSid As Long)
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As Any, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function IsValidSid Lib "advapi32.dll" (ByVal pSid As Long) As Long
Private Declare Function GetSidSubAuthority Lib "advapi32.dll" (ByVal pSid As Long, ByVal nSubAuthority As Long) As Long
Private Declare Function GetSidSubAuthorityCount Lib "advapi32.dll" (ByVal pSid As Long) As Long
Private Declare Function EqualSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal lSize As Long)
'Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueW" (ByVal lpSystemName As Long, ByVal lpName As Long, lpLuid As Long) As Long
'Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
'Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long
'Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long

Private Const TOKEN_QUERY               As Long = &H8&
Private Const TokenElevation            As Long = 20&

Function IsProcessElevated(Optional hProcess As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim hToken           As Long
    Dim dwLengthNeeded   As Long
    Dim dwIsElevated     As Long
    Dim osi              As RTL_OSVERSIONINFOEXW
    
    osi.dwOSVersionInfoSize = Len(osi)
    RtlGetVersion osi
    
    ' < Win Vista. Устанавливаем true, если пользователь состоит в группе "Администраторы"
    If osi.dwMajorVersion < 6 Then IsProcessElevated = (GetUserType() = "Administrator"): Exit Function

    If hProcess = 0 Then hProcess = GetCurrentProcess()
    
    If OpenProcessToken(hProcess, TOKEN_QUERY, hToken) Then

        If 0 <> GetTokenInformation(hToken, TokenElevation, dwIsElevated, 4&, dwLengthNeeded) Then
            IsProcessElevated = (dwIsElevated <> 0)
        End If
        
        CloseHandle hToken: hToken = 0
    End If
    
    Exit Function
ErrorHandler:
    WriteC "clsOSver.IsProcessElevated", cErr
    If hToken Then CloseHandle hToken: hToken = 0
End Function

Public Function GetUserType(Optional hProcess As Long) As String
    On Error GoTo ErrorHandler

    Const TOKEN_QUERY                   As Long = &H8&
    Const SECURITY_NT_AUTHORITY         As Long = 5&
    Const TokenGroups                   As Long = 2&
    Const SECURITY_BUILTIN_DOMAIN_RID   As Long = &H20&
    Const DOMAIN_ALIAS_RID_ADMINS       As Long = &H220&
    Const DOMAIN_ALIAS_RID_USERS        As Long = &H221&
    Const DOMAIN_ALIAS_RID_GUESTS       As Long = &H222&
    Const DOMAIN_ALIAS_RID_POWER_USERS  As Long = &H223&

    Dim hProcessToken   As Long
    Dim BufferSize      As Long
    Dim psidAdmin       As Long
    Dim psidPower       As Long
    Dim psidUser        As Long
    Dim psidGuest       As Long
    Dim lResult         As Long
    Dim i               As Long
    Dim tpTokens        As TOKEN_GROUPS
    Dim tpSidAuth       As SID_IDENTIFIER_AUTHORITY
    
    GetUserType = "Unknown"
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
    
    ' в идеале, сначала нужно проверять токен, полученный от потока
    ' If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) Then
    ' ограничимся токеном процесса, т.к. пока не планируем более 1 потока
    
    If hProcess = 0 Then hProcess = GetCurrentProcess()
    If 0 = OpenProcessToken(hProcess, TOKEN_QUERY, hProcessToken) Then Exit Function
    
    If hProcessToken Then

        ' Определяем требуемый размер буфера
        GetTokenInformation hProcessToken, ByVal TokenGroups, 0&, 0&, BufferSize
        
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long  ' Переводим размер byte -> Long
            
            ' Получаем информацию о SID-ах групп, ассоциированных с этим токеном
            If 0 <> GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize) Then
            
                ' Заполняем структуру из буфера
                Call CopyMemory(tpTokens, InfoBuffer(0), Len(tpTokens))
            
                ' Получаем SID-ы каждого типа пользователей
                lResult = AllocateAndInitializeSid(tpSidAuth, 2&, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0&, 0&, 0&, 0&, 0&, 0&, psidAdmin)
                lResult = AllocateAndInitializeSid(tpSidAuth, 2&, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_POWER_USERS, 0&, 0&, 0&, 0&, 0&, 0&, psidPower)
                lResult = AllocateAndInitializeSid(tpSidAuth, 2&, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_USERS, 0&, 0&, 0&, 0&, 0&, 0&, psidUser)
                lResult = AllocateAndInitializeSid(tpSidAuth, 2&, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_GUESTS, 0&, 0&, 0&, 0&, 0&, 0&, psidGuest)
            
                If IsValidSid(psidAdmin) And IsValidSid(psidPower) And IsValidSid(psidUser) And IsValidSid(psidGuest) Then
                  
                    For i = 0 To tpTokens.GroupCount
                        ' Берем SID каждой из ассоциированных групп
                        If IsValidSid(tpTokens.Groups(i).SID) Then
                            ' Проверяем на соответствие
                            If EqualSid(ByVal tpTokens.Groups(i).SID, ByVal psidAdmin) Then
                                GetUserType = "Administrator":  Exit For
                            ElseIf EqualSid(ByVal tpTokens.Groups(i).SID, ByVal psidPower) Then
                                GetUserType = "Power User":     Exit For
                            ElseIf EqualSid(ByVal tpTokens.Groups(i).SID, ByVal psidUser) Then
                                GetUserType = "Limited User":   Exit For
                            ElseIf EqualSid(ByVal tpTokens.Groups(i).SID, ByVal psidGuest) Then
                                GetUserType = "Guest":          Exit For
                            End If
                        End If
                    Next
                End If
                If psidAdmin Then FreeSid psidAdmin
                If psidPower Then FreeSid psidPower
                If psidUser Then FreeSid psidUser
                If psidGuest Then FreeSid psidGuest
            End If
        End If
        CloseHandle hProcessToken: hProcessToken = 0
    End If
    
    Exit Function
ErrorHandler:
    WriteC "clsOSver.GetUserType", cErr
    If hProcessToken Then CloseHandle hProcessToken: hProcessToken = 0
    If psidAdmin Then FreeSid psidAdmin
    If psidPower Then FreeSid psidPower
    If psidUser Then FreeSid psidUser
    If psidGuest Then FreeSid psidGuest
End Function

