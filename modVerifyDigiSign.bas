Attribute VB_Name = "modVerifyDigiSign"
Option Explicit

'
' Authenticode digital signature verifier / Driver's verifier for compliance with WHQL standard
'
' Copyrights: (с) Pol'shyn Stanislav Viktorovich aka Alex Dragokas
'
' Also get information about:
' - Signer
' - TimeStamp
' - Signature hash (look at module 'modSignerInfo')

'Some code examples:
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa382384%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
'https://support.microsoft.com/en-us/kb/323809?wa=wsignin1.0
'https://msdn.microsoft.com/en-us/library/aa382384.aspx
'http://rsdn.ru/forum/src/3152752.hot
'http://rsdn.ru/forum/winapi/2731079.hot
'http://processhacker.sourceforge.net/doc/verify_8c_source.html
'http://eternalwindows.jp/crypto/certverify/certverify03.html
'http://forum.sysinternals.com/howto-verify-the-digital-signature-of-a-file_topic19247.html
'http://processhacker.sourceforge.net/doc/verify_8c_source.html

'CERT_CHAIN_POLICY_STATUS structure
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa377188%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396

'Error information:
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa378137(v=vs.85).aspx

' revision 2.2. (17.05.2016)
' Added SHA256 support

#Const UseHashTable = True ' использовать хеш-таблицы? (Maded by Кривоус Анатолий)

Const MAX_PATH As Long = 260&

Public Type SignResult_TYPE ' Digital signature data
    isSigned     As Boolean ' signed?
    isLegit      As Boolean ' is signature legitimate?
    isCert       As Boolean ' is signed by Windows security catalogue?
    Issuer       As String  ' signer name
    RootCertHash As String
    ShortMessage As String  ' short description of checking results
    FullMessage  As String  ' full description of cheking results
    ReturnCode   As String  ' result error code of WinVerifyTrust
End Type

Public Enum FLAGS_SignVerify
    SV_CheckHoleChain = 1       ' - check whole trust chain ( require internet connection )
    SV_DoNotUseHashChecking = 2 ' - do not use system cache ( checking by hash )
    SV_DisableCatalogVerify = 4 ' - do not use checking by security catalogue
    SV_isDriver = 8             ' - verify driver for compliance with WHQL standard
    SV_AllowSelfSigned = 16     ' - self-signed certificates should be considered as legitimate
End Enum

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type CATALOG_INFO
    cbStruct As Long
    wszCatalogFile(MAX_PATH - 1) As Integer
End Type

Private Type WINTRUST_FILE_INFO
    cbStruct As Long
    pcwszFilePath As Long
    hFile As Long
    pgKnownSubject As Long
End Type

Private Type WINTRUST_CATALOG_INFO
    cbStruct As Long
    dwCatalogVersion As Long
    pcwszCatalogFilePath As Long
    pcwszMemberTag As Long
    pcwszMemberFilePath As Long
    hMemberFile As Long
    pbCalculatedFileHash As Long
    cbCalculatedFileHash As Long
    pcCatalogContext As Long
    hCatAdmin As Long
End Type

Private Type WINTRUST_DATA
    cbStruct As Long
    pPolicyCallbackData As Long
    pSIPClientData As Long
    dwUIChoice As Long
    fdwRevocationChecks As Long
    dwUnionChoice As Long
    pUnion As Long                      'ptr to one of 5 structures based on dwUnionChoice param
    dwStateAction As Long
    hWVTStateData As Long
    pwszURLReference As Long
    dwProvFlags As Long
    dwUIContext As Long
    pSignatureSettings As Long          'ptr to WINTRUST_SIGNATURE_SETTINGS (Win8+)
End Type

Private Type CERT_STRONG_SIGN_PARA
    cbSize As Long
    dwInfoChoice As Long
    pszOID As Long
End Type

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszGuid As Long, pGuid As GUID) As Long
Private Declare Function CryptCATAdminAcquireContext Lib "Wintrust.dll" (hCatAdmin As Long, ByVal pgSubsystem As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCATAdminAcquireContext2 Lib "Wintrust.dll" (hCatAdmin As Long, ByVal pgSubsystem As Long, ByVal pwszHashAlgorithm As Long, ByVal pStrongHashPolicy As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCATAdminReleaseContext Lib "Wintrust.dll" (ByVal hCatAdmin As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCATAdminCalcHashFromFileHandle Lib "Wintrust.dll" (ByVal hFile As Long, pcbHash As Long, pbHash As Byte, ByVal dwFlags As Long) As Long
Private Declare Function CryptCATAdminCalcHashFromFileHandle2 Lib "Wintrust.dll" (ByVal hCatAdmin As Long, ByVal hFile As Long, pcbHash As Long, pbHash As Byte, ByVal dwFlags As Long) As Long
Private Declare Function CryptCATAdminEnumCatalogFromHash Lib "Wintrust.dll" (ByVal hCatAdmin As Long, pbHash As Byte, ByVal cbHash As Long, ByVal dwFlags As Long, phPrevCatInfo As Long) As Long
Private Declare Function CryptCATCatalogInfoFromContext Lib "Wintrust.dll" (ByVal hCatInfo As Long, psCatInfo As CATALOG_INFO, ByVal dwFlags As Long) As Long
Private Declare Function CryptCATAdminReleaseCatalogContext Lib "Wintrust.dll" (ByVal hCatAdmin As Long, ByVal hCatInfo As Long, ByVal dwFlags As Long) As Long
Private Declare Function WinVerifyTrust Lib "Wintrust.dll" (ByVal hWnd As Long, pgActionID As GUID, ByVal pWVTData As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Arguments As Any) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long

'Private Declare Function WinVerifyFile Lib "istrusted.dll" Alias "Checkfile" (ByVal sFilename As String) As Boolean

'Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
'Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (byval OldValue As Long) As Long

Const WTD_UI_NONE                   As Long = 2&
' checking certificate revocation
Const WTD_REVOKE_NONE               As Long = 0&    ' do not check certificates for revocation
Const WTD_REVOKE_WHOLECHAIN         As Long = 1&    ' check for certificate revocation all chain
' method of verification
Const WTD_CHOICE_CATALOG            As Long = 2&    ' check by certificate that is stored in local windows security storage
Const WTD_CHOICE_FILE               As Long = 1&    ' check internal signature
' flags
Const WTD_SAFER_FLAG                As Long = 256&   ' ??? (probably, no UI for XP SP2)
Const WTD_REVOCATION_CHECK_NONE     As Long = 16&    ' do not execute revecation checking of cert. chain
Const WTD_REVOCATION_CHECK_END_CERT As Long = &H20&  ' check for revocation end cert. only
Const WTD_REVOCATION_CHECK_CHAIN    As Long = &H40&  ' check all cert. chain ( require internet connection to port 53 TCP/UDP )
Const WTD_REVOCATION_CHECK_CHAIN_EXCLUDE_ROOT As Long = &H80& ' check all chain, excepting the root cert.
Const WTD_HASH_ONLY_FLAG            As Long = &H200& ' checking by hash only
Const WTD_NO_POLICY_USAGE_FLAG      As Long = 4&     ' do not mention on local security policy settings
Const WTD_CACHE_ONLY_URL_RETRIEVAL  As Long = 4096&  ' check for certificate revocation but only by local cache
Const WTD_LIFETIME_SIGNING_FLAG     As Long = &H800& ' check exparation date of certificate
' action
Const WTD_STATEACTION_VERIFY        As Long = 1&
Const WTD_STATEACTION_IGNORE        As Long = 0&
Const WTD_STATEACTION_CLOSE         As Long = 2&
' context
Const WTD_UICONTEXT_EXECUTE         As Long = 0&
' errors
Const TRUST_E_SUBJECT_NOT_TRUSTED   As Long = &H800B0004
Const TRUST_E_PROVIDER_UNKNOWN      As Long = &H800B0001
Const TRUST_E_ACTION_UNKNOWN        As Long = &H800B0002
Const TRUST_E_SUBJECT_FORM_UNKNOWN  As Long = &H800B0003
Const CERT_E_REVOKED                As Long = &H800B010C
Const CERT_E_EXPIRED                As Long = &H800B0101
Const TRUST_E_BAD_DIGEST            As Long = &H80096010
Const TRUST_E_NOSIGNATURE           As Long = &H800B0100
Const TRUST_E_EXPLICIT_DISTRUST     As Long = &H800B0111
Const CRYPT_E_SECURITY_SETTINGS     As Long = &H80092026
Const CERT_E_UNTRUSTEDROOT          As Long = &H800B0109
Const CERT_E_PURPOSE                As Long = &H800B0106
'OID
Const szOID_CERT_STRONG_SIGN_OS_1   As String = "1.3.6.1.4.1.311.72.1.1"
Const szOID_CERT_STRONG_KEY_OS_1    As String = "1.3.6.1.4.1.311.72.2.1"
'Crypt Algorithms
Const BCRYPT_SHA1_ALGORITHM         As String = "SHA1"      '160-bit
Const BCRYPT_SHA256_ALGORITHM       As String = "SHA256"    '256-bit

Const CERT_STRONG_SIGN_OID_INFO_CHOICE As Long = 2&

' other
Const INVALID_HANDLE_VALUE          As Long = -1&
Const ERROR_INSUFFICIENT_BUFFER     As Long = 122&

Dim WINTRUST_ACTION_GENERIC_VERIFY_V2   As GUID
Dim DRIVER_ACTION_VERIFY                As GUID

Public cOut         As Long
Public cErr         As Long


Public Function SignVerify( _
    FilePath As String, _
    flags As FLAGS_SignVerify, _
    SignResult As SignResult_TYPE) As Boolean
 
    On Error GoTo ErrorHandler
 
    ' in.  FilePath - path to PE EXE file for validation
    ' in.  Flags - options for checking
    ' out. SignResult struct
    
    ' RETURN value - return true, if the integrity of the executable file is confirmed, notwithstanding:
    ' - possible restrictions in the local policy settings
    ' - self-signed certificate type (if the option 'CheckHoleChain = true' is not specified and revocation data are not cached)
    ' - checking for certificate exparation is not performed. If needed, add a flag WTD_LIFETIME_SIGNING_FLAG
    
    ' For even more strong verification (forbid reading revocation info from the cache),
    ' replace the flag WTD_CACHE_ONLY_URL_RETRIEVAL into WTD_REVOCATION_CHECK_NONE.
    ' Note that certificate revocation is a specific procedure and it should be performed
    ' only if you suspect that digital signature has been stolen or used in malware
    ' (this kind of verification require internet connection, can freeze a program and time-consuming).
    
    ' in. Flags (can be combined by 'OR' statement):
    
    ' SV_CheckHoleChain       ' - check whole trust chain ( require internet connection )
    ' SV_DoNotUseHashChecking ' - do not use system cache ( checking by hash )
    ' SV_DisableCatalogVerify ' - do not use checking by security catalogue
    ' SV_isDriver             ' - verify driver for compliance with WHQL standard
    ' SV_AllowSelfSigned      ' - self-signed certificates should be considered as legitimate
    
    Dim InfoStruct               As CATALOG_INFO
    Dim WintrustStructure        As WINTRUST_DATA
    Dim WintrustCatalogStructure As WINTRUST_CATALOG_INFO
    Dim WintrustFileStructure    As WINTRUST_FILE_INFO
    Dim CertSignPara             As CERT_STRONG_SIGN_PARA
    
    Static MajorMinor       As Single
    Static IsVistaAndLater  As Boolean
    Static IsWin8AndLater   As Boolean
    Static SignCache()      As SignResult_TYPE
    Static SC_pos           As Long
    #If UseHashTable Then
        Static oSignIndex As clsTrickHashTable
    #Else
        Static oSignIndex As Object
    #End If
    
    Dim i               As Long
    Dim Context         As Long
    Dim FileHandle      As Long
    Dim HashSize        As Long
    Dim aBuf()          As Byte
    Dim sBuf            As String
    Dim CatalogContext  As Long
    Dim MemberTag       As String
    Dim ReturnFlag      As Boolean
    Dim hLib            As Long
    Dim ReturnVal       As Long
    Dim inf(68)         As Long
    Dim ActionGuid      As GUID
    Dim Success         As Boolean
    
    With SignResult     'clear results of checking
        .ReturnCode = 0
        .FullMessage = vbNullString
        .ShortMessage = vbNullString
        .Issuer = vbNullString
        .RootCertHash = vbNullString
        .isSigned = False
        .isLegit = False
        .isCert = False
    End With
    
    ToggleWow64FSRedirection True
    
    If 0 = ObjPtr(oSignIndex) Then
        #If UseHashTable Then
            Set oSignIndex = New clsTrickHashTable
        #Else
            Set oSignIndex = CreateObject("Scripting.Dictionary")
        #End If
        oSignIndex.CompareMode = vbTextCompare
        ReDim SignCache(100)
    Else
        If oSignIndex.Exists(FilePath) Then
            SignResult = SignCache(oSignIndex(FilePath))
            Exit Function
        End If
    End If
    
    SC_pos = SC_pos + 1
    If UBound(SignCache) < SC_pos Then ReDim Preserve SignCache(UBound(SignCache) + 100)
    oSignIndex.Add FilePath, SC_pos
    
    If MajorMinor = 0 Then  'not cached
        hLib = LoadLibrary(StrPtr("Wintrust.dll"))
        If hLib = 0 Then
            WriteCon "NOT SUPPORTED. LastDllErr = 0x" & Hex(err.LastDllError)
            SignResult.ShortMessage = "NOT SUPPORTED."
            SignCache(SC_pos) = SignResult
            Exit Function
        End If
    Else
        FreeLibrary hLib
    End If
    
    If MajorMinor = 0 Then
        inf(0) = 276: GetVersionEx inf(0): IsVistaAndLater = (inf(1) >= 6): MajorMinor = inf(1) + inf(2) / 10: IsWin8AndLater = (MajorMinor >= 6.2)
    End If
    
    CLSIDFromString StrPtr("{F750E6C3-38EE-11D1-85E5-00C04FC295EE}"), DRIVER_ACTION_VERIFY
    CLSIDFromString StrPtr("{00AAC56B-CD44-11D0-8CC2-00C04FC295EE}"), WINTRUST_ACTION_GENERIC_VERIFY_V2
    
    If (flags And SV_isDriver) Then
        ActionGuid = DRIVER_ACTION_VERIFY
    Else
        ActionGuid = WINTRUST_ACTION_GENERIC_VERIFY_V2
    End If
    
    InfoStruct.cbStruct = Len(InfoStruct)
    WintrustFileStructure.cbStruct = Len(WintrustFileStructure)
    
'    With CertSignPara
'        .cbSize = LenB(CertSignPara)
'        .dwInfoChoice = CERT_STRONG_SIGN_OID_INFO_CHOICE
'        'szOID_CERT_STRONG_SIGN_OS_1    'SHA2
'        'szOID_CERT_STRONG_KEY_OS_1     'SHA1 + SHA2
'        .pszOID = StrPtr(szOID_CERT_STRONG_SIGN_OS_1)
'    End With
    
    'VarPtr(CertSignPara)
    'StrPtr(BCRYPT_SHA256_ALGORITHM)
    
    If MajorMinor >= 6.2 Then
        ' obtain context for procedure of signature verification
        CryptCATAdminAcquireContext2 Context, VarPtr(DRIVER_ACTION_VERIFY), StrPtr(BCRYPT_SHA256_ALGORITHM), 0&, 0&
    End If
    
    If Context = 0 Then
        If Not (CBool(CryptCATAdminAcquireContext(Context, VarPtr(DRIVER_ACTION_VERIFY), 0&))) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminAcquireContext", FilePath
            SignCache(SC_pos) = SignResult
            Exit Function
        End If
    End If
    
    If Not FileExists(FilePath) Then
        CryptCATAdminReleaseContext Context, 0&
        SignCache(SC_pos) = SignResult
        Exit Function
    End If
    
    ' opening the file
    ToggleWow64FSRedirection False, FilePath
    
    OpenW FilePath, FOR_READ, FileHandle
    
    ToggleWow64FSRedirection True
    
    If (INVALID_HANDLE_VALUE = FileHandle) Then
        WriteError err, SignResult.ShortMessage, "CreateFile", FilePath
        CryptCATAdminReleaseContext Context, 0&
        SignCache(SC_pos) = SignResult
        Exit Function
    End If
    
    ' размер файла = 0
    If LOFW(FileHandle) = 0 Then
        CryptCATAdminReleaseContext Context, 0&
        SignCache(SC_pos) = SignResult
        CloseHandle FileHandle
        Exit Function
    End If
    
    ' checking, is internal signature present
    'SignResult.isSigned = IsSignPresent(FilePath) ', FileHandle)
    
    If MajorMinor >= 6.2 Then
        ' obtain size needed for hash
        HashSize = 1
        ReDim aBuf(HashSize - 1&)
        Success = CryptCATAdminCalcHashFromFileHandle2(Context, FileHandle, HashSize, ByVal 0&, 0&)
        
        If err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
            Success = False
            ReDim aBuf(HashSize - 1&)
            If CBool(CryptCATAdminCalcHashFromFileHandle2(Context, FileHandle, HashSize, aBuf(0), 0&)) Then Success = True
        End If
    End If
    
    If (HashSize = 0& Or Not Success) Then
        ' obtain size needed for hash
        CryptCATAdminCalcHashFromFileHandle FileHandle, HashSize, ByVal 0&, 0&
    
        If (HashSize = 0&) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminCalcHashFromFileHandle", FilePath
            CryptCATAdminReleaseContext Context, 0&
            SignCache(SC_pos) = SignResult
            CloseHandle FileHandle: FileHandle = 0
            Exit Function
        End If

        ' allocating the memory
        ReDim aBuf(HashSize - 1&)

        ' calculation of the hash
        If Not CBool(CryptCATAdminCalcHashFromFileHandle(FileHandle, HashSize, aBuf(0), 0&)) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminCalcHashFromFileHandle", FilePath
            CryptCATAdminReleaseContext Context, 0&
            SignCache(SC_pos) = SignResult
            CloseHandle FileHandle: FileHandle = 0
            Exit Function
        End If
    End If
    
    ' Converting hash into string
    For i = 0& To UBound(aBuf)
        MemberTag = MemberTag & Right$("0" & Hex(aBuf(i)), 2&)
    Next

    'here is M$ bug with  C:\WINDOWS\system32\catroot2\{GUID} file. // Now we handle it.

    ' Obtain catalogue for our context
    If Not HasCatRootVulnerability() Then
        CatalogContext = CryptCATAdminEnumCatalogFromHash(Context, aBuf(0), HashSize, 0&, ByVal 0&)
    End If

    If (CatalogContext) Then
        ' If unable to get information
        If Not CBool(CryptCATCatalogInfoFromContext(CatalogContext, InfoStruct, 0&)) Then
            WriteError err, SignResult.ShortMessage, "CryptCATCatalogInfoFromContext", FilePath
            ' Release context
            CryptCATAdminReleaseCatalogContext Context, CatalogContext, 0&
            SignCache(SC_pos) = SignResult
            CatalogContext = 0&
        End If
        
        'sBuf = String$(MAX_PATH, vbNullChar)
        'memcpy ByVal StrPtr(sBuf), InfoStruct.wszCatalogFile(0), MAX_PATH
        'sBuf = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
        'Debug.Print "Hash: " & sBuf
    End If
    
    ' If we got a valid context, verify the signature through the catalog.
    ' Otherwise (if Embedded signature is exist or flag "Ignore checking by catalogue" is set), trying to verify internal signature of the file:
    If (CatalogContext = 0& Or (flags And SV_DisableCatalogVerify)) Or SignResult.isSigned Then
        With WintrustFileStructure                  'WINTRUST_FILE_INFO
            .cbStruct = Len(WintrustFileStructure)
            .pcwszFilePath = StrPtr(FilePath)
            .hFile = 0&
            '.pgKnownSubject = 0
        End With
        
        With WintrustStructure                      'WINTRUST_DATA
            .cbStruct = Len(WintrustStructure)
            .dwUnionChoice = WTD_CHOICE_FILE
            .pUnion = VarPtr(WintrustFileStructure)     'pFile
            .dwUIChoice = WTD_UI_NONE
            .dwStateAction = WTD_STATEACTION_VERIFY 'WTD_STATEACTION_IGNORE
            .hWVTStateData = 0&
            .pwszURLReference = 0&
            ' ---------------------------------------- Перечень флагов -------------------------------------------
            .fdwRevocationChecks = IIf(flags And SV_CheckHoleChain, WTD_REVOKE_WHOLECHAIN, WTD_REVOKE_NONE)   ' checking for cert. revokation
            If flags And SV_CheckHoleChain Then
                .dwProvFlags = .dwProvFlags Or WTD_REVOCATION_CHECK_CHAIN
            Else
                ' take data about cert. chain verification from local cache only, if they were saved ( >= Vista ). Do not use internet connection.
                .dwProvFlags = .dwProvFlags Or IIf(IsVistaAndLater, WTD_CACHE_ONLY_URL_RETRIEVAL, WTD_REVOCATION_CHECK_NONE)
            End If
            '.dwProvFlags = .dwProvFlags Or WTD_NO_POLICY_USAGE_FLAG                                          ' do not check certificate purpose (disabled)
            .dwProvFlags = .dwProvFlags Or WTD_LIFETIME_SIGNING_FLAG                                          ' invalidate expired signatures
            If 0 = (flags And SV_DoNotUseHashChecking) Then .dwProvFlags = .dwProvFlags Or WTD_HASH_ONLY_FLAG ' check only by hash
            .dwProvFlags = .dwProvFlags Or WTD_SAFER_FLAG                                                     ' without UI
        End With
    
    Else
        ' We received information from the catalogue. Check validity through it.
        SignResult.isSigned = True
        SignResult.isCert = True
        With WintrustStructure                      'WINTRUST_DATA
            .cbStruct = Len(WintrustStructure)
            .pPolicyCallbackData = 0&
            .pSIPClientData = 0&
            .dwUIChoice = WTD_UI_NONE
            .fdwRevocationChecks = WTD_REVOKE_NONE
            .dwUnionChoice = WTD_CHOICE_CATALOG
            .pUnion = VarPtr(WintrustCatalogStructure)   'pCatalog
            .dwStateAction = WTD_STATEACTION_VERIFY
            .hWVTStateData = 0&
            .pwszURLReference = 0&
            .dwProvFlags = WTD_SAFER_FLAG
            If Not CBool(flags And SV_CheckHoleChain) Then
                .dwProvFlags = .dwProvFlags Or IIf(IsVistaAndLater, WTD_CACHE_ONLY_URL_RETRIEVAL, WTD_REVOCATION_CHECK_NONE)
            End If
            .dwUIContext = WTD_UICONTEXT_EXECUTE
        End With
        
        ' Fill in catalogue structure
        With WintrustCatalogStructure               'WINTRUST_CATALOG_INFO
            .cbStruct = Len(WintrustCatalogStructure)
            .dwCatalogVersion = 0&
            .pcwszCatalogFilePath = VarPtr(InfoStruct.wszCatalogFile(0))
            .pcwszMemberTag = StrPtr(MemberTag)
            .pcwszMemberFilePath = StrPtr(FilePath)
            .hMemberFile = 0&
            '.cbCalculatedFileHash = HashSize
            '.pbCalculatedFileHash = VarPtr(aBuf(0))
        End With
    End If
    
    ToggleWow64FSRedirection False, FilePath
    
    ' calling main verification function
    ReturnVal = WinVerifyTrust(INVALID_HANDLE_VALUE, ActionGuid, VarPtr(WintrustStructure)) ' INVALID_HANDLE_VALUE mean non-interactive checking (without UI)
    'Stop
    ToggleWow64FSRedirection True

    'Debug.Print WintrustStructure.hWVTStateData    ' -> identifier, in which individual signatures can be extracted

    GetSignerInfo WintrustStructure.hWVTStateData, SignResult.Issuer, SignResult.RootCertHash

    ' SV_AllowSelfSigned - allow self-signed certificate ( ignore lack of confidence in the final certificate )
    
    If flags And SV_AllowSelfSigned Then
        ReturnFlag = ((ReturnVal = 0) Or (ReturnVal = CERT_E_UNTRUSTEDROOT))
    Else
        ReturnFlag = (ReturnVal = 0)
    End If
    
    With SignResult
    
        .isSigned = True
    
        If True = ReturnFlag Then .isSigned = True

        Select Case ReturnVal
        Case 0
            .ShortMessage = "Legit signature."
        Case TRUST_E_SUBJECT_NOT_TRUSTED
            .ShortMessage = "TRUST_E_SUBJECT_NOT_TRUSTED"
            'The user clicked "No" when asked to install and run.
        Case TRUST_E_PROVIDER_UNKNOWN
            .ShortMessage = "TRUST_E_PROVIDER_UNKNOWN"
            'The trust provider is not recognized on this system.
        Case TRUST_E_ACTION_UNKNOWN
            .ShortMessage = "TRUST_E_ACTION_UNKNOWN"
            'The trust provider does not support the specified action.
        Case TRUST_E_SUBJECT_FORM_UNKNOWN
            .ShortMessage = "TRUST_E_SUBJECT_FORM_UNKNOWN"
            'This can happen when WinVerifyTrust is called on an unknown file type
        Case CERT_E_REVOKED
            .ShortMessage = "CERT_E_REVOKED"
            'A certificate was explicitly revoked by its issuer.
        Case CERT_E_EXPIRED
            .ShortMessage = "CERT_E_EXPIRED"
            'A required certificate is not within its validity period when verifying against the current system clock or the timestamp in the signed file
        Case CERT_E_PURPOSE
            .ShortMessage = "CERT_E_PURPOSE"
            'The certificate is being used for a purpose other than one specified by the issuing CA.
        Case TRUST_E_BAD_DIGEST
            .ShortMessage = "TRUST_E_BAD_DIGEST"
            'This will happen if the file has been modified or corruped.
        Case TRUST_E_NOSIGNATURE
            If TRUST_E_NOSIGNATURE = err.LastDllError Or _
                TRUST_E_SUBJECT_FORM_UNKNOWN = err.LastDllError Or _
                TRUST_E_PROVIDER_UNKNOWN = err.LastDllError Or _
                err.LastDllError = 0 Then
                .ShortMessage = "TRUST_E_NOSIGNATURE: Not signed"
                .isSigned = False
            Else
                .ShortMessage = "TRUST_E_NOSIGNATURE: Not valid signature"
                'The signature was not valid or there was an error opening the file.
            End If
        Case TRUST_E_EXPLICIT_DISTRUST
            .ShortMessage = "TRUST_E_EXPLICIT_DISTRUST: Signature is forbidden"
            'The signature Is present, but specifically disallowed
            'The hash that represents the subject or the publisher is not allowed by the admin or user.
        Case CRYPT_E_SECURITY_SETTINGS
            .ShortMessage = "CRYPT_E_SECURITY_SETTINGS"
            ' The hash that represents the subject or the publisher was not explicitly trusted by the admin and the
            ' admin policy has disabled user trust. No signature, publisher or time stamp errors.
        Case CERT_E_UNTRUSTEDROOT
            .ShortMessage = "CERT_E_UNTRUSTEDROOT: Verified, but self-signed"
            'A certificate chain processed, but terminated in a root certificate which is not trusted by the trust provider.
        Case Else
            .ShortMessage = "Other error. Code = " & ReturnVal & ". LastDLLError = " & err.LastDllError
            'The UI was disabled in dwUIChoice or the admin policy has disabled user trust. ReturnVal contains the publisher or time stamp chain error.
        End Select
    
        ' Other error codes can be found on MSDN:
        ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa377188%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
        ' https://msdn.microsoft.com/en-us/library/ee488436.aspx
        ' This is not an exhaustive list.
    
        .FullMessage = ErrMessageText(ReturnVal)
    
        If .FullMessage <> "Операция успешно завершена." And .FullMessage <> "В этом объекте нет подписи." Then
            Debug.Print FilePath & vbTab & vbTab & .FullMessage
        End If
    
        .ReturnCode = ReturnVal
        .isLegit = ReturnFlag
        SignVerify = .isLegit
    
    End With
    
    ' Release context
    If (CatalogContext) Then
        CryptCATAdminReleaseCatalogContext Context, CatalogContext, 0&
    End If
    
    ' Free memory, used by provider during signature verification
    WintrustStructure.dwStateAction = WTD_STATEACTION_CLOSE
    WinVerifyTrust INVALID_HANDLE_VALUE, ActionGuid, VarPtr(WintrustStructure)
    
    ' closing the file, release context
    CloseHandle FileHandle: FileHandle = 0
    CryptCATAdminReleaseContext Context, 0&
    
    SignCache(SC_pos) = SignResult
    
    Exit Function
ErrorHandler:
    ErrorMsg err, "SignVerify", FilePath
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Function

Function HasCatRootVulnerability() As Boolean
    On Error GoTo ErrHandler
    Static IsInit       As Boolean
    Static VulnStatus   As Boolean
    
    If IsInit Then
        HasCatRootVulnerability = VulnStatus
        Exit Function
    Else
        IsInit = True
    End If
    
    Dim inf(68) As Long: inf(0) = 276: GetVersionEx inf(0): If inf(1) < 6 Then Exit Function
    
    Dim sFile   As String
    Dim lr      As Long
    Dim WinDir  As String
    
    WinDir = Space$(MAX_PATH)
    lr = GetWindowsDirectory(StrPtr(WinDir), MAX_PATH)
    If lr Then WinDir = Left$(WinDir, lr)
    sFile = Dir$(WinDir & "\System32\catroot2\*") 'not affected by wow64
    Do While Len(sFile)
        If sFile Like "{????????????????????????????????????}" Then
            lr = GetFileAttributes(StrPtr(WinDir & "\System32\catroot2\" & sFile))
            If lr <> INVALID_HANDLE_VALUE And (lr And vbDirectory) Then
                VulnStatus = True: HasCatRootVulnerability = True: Exit Function
            End If
        End If
        sFile = Dir$()
    Loop
    Exit Function
ErrHandler:
    ErrorMsg err, "HasCatRootVulnerability"
    If inIDE Then Stop: Resume Next
End Function

Public Function ErrMessageText(lCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
    Const FORMAT_MESSAGE_FROM_HMODULE   As Long = &H800&
    
    Dim sRtrnMessage   As String
    Dim lret           As Long
    
    sRtrnMessage = String$(MAX_PATH, vbNullChar)
    lret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lCode, 0&, StrPtr(sRtrnMessage), MAX_PATH, ByVal 0&)
    If lret > 0 Then
        ErrMessageText = Left$(sRtrnMessage, lret)
        ErrMessageText = Replace$(ErrMessageText, vbCrLf, vbNullString)
    End If
End Function

' Will write error with description to console and in ShortMessage (RetErrMessage)
Function WriteError(err As ErrObject, RetErrMessage As String, FunctionName As String, FileName As String)
    Dim ErrNumber As Long
    ErrNumber = err.LastDllError
    
    If &H800700C1 = ErrNumber Then
        ' if we got "%1 is not a valid Win32 application." and PE EXE contain pointer to SecurityDir struct,
        ' it's mean digital signature was damaged
        ' https://chentiangemalc.wordpress.com/2014/08/01/case-of-the-server-returned-a-referral/

        If IsSignPresent(FileName) Then
            'Signature is damaged!
            RetErrMessage = Translate(1862)
        Else
            RetErrMessage = ErrMessageText(ErrNumber)
        End If
    Else
        RetErrMessage = ErrMessageText(ErrNumber)
    End If
    
    WriteCon "Error in " & FunctionName & ": " & "0x" & Hex(ErrNumber) & ". " & RetErrMessage & ". File: " & FileName
End Function

Public Sub WriteCon(ByVal txt As String, Optional cHandle As Long, Optional NoNewLine As Boolean = False)
    Debug.Print txt
    Exit Sub
    
    Const CP_ACP    As Long = 0&
    Const CP_OEMCP  As Long = 1&
    If cHandle = 0 Then cHandle = cOut
    Debug.Print txt
    If cHandle <> 0 Then
        If Not NoNewLine Then txt = txt & vbNewLine
        'txt = ConvertCodePage(txt, CP_ACP, CP_OEMCP)
        'WriteConsole cHandle, ByVal txt, Len(txt), 0, ByVal 0&
    End If
End Sub


' ---------------------------------------------------------------------------------------------------
' StartupList2 routine
' ---------------------------------------------------------------------------------------------------

Public Function VerifyFileSignature(sFile$) As Integer
'    If Not FileExists(App.Path & "\istrusted.dll") Then
'        If msgboxw("To verify file signatures, StartupList needs to " & _
'                  "download an external library from www.merijn.org. " & _
'                  vbCrLf & vbCrLf & "Continue?", vbYesNo + vbQuestion) = vbYes Then
'            If DownloadFile("http://www.merijn.org/files/istrusted.dll", App.Path & "\istrusted.dll") Then
'                'file downloaded ok, continue
'            Else
'                'file download failed
'                bAbort = True
'                VerifyFileSignature = -1
'                Exit Function
'            End If
'        Else
'            'user aborted download
'            bAbort = True
'            VerifyFileSignature = -1
'            Exit Function
'        End If
'    End If
    
    If WinVerifyFile(sFile) Then
        VerifyFileSignature = 1
    Else
        VerifyFileSignature = 0
    End If
End Function

Public Sub WinTrustVerifyChildNodes(sKey$)
    If bAbort Then Exit Sub
    If Not NodeExists(sKey) Then Exit Sub
    Dim nodFirst As Node, nodCurr As Node
    Set nodFirst = frmStartupList2.tvwMain.Nodes(sKey).Child
    Set nodCurr = nodFirst
    If Not (nodCurr Is Nothing) Then
        Do
            If nodCurr.Children > 0 Then WinTrustVerifyChildNodes nodCurr.Key
        
            WinTrustVerifyNode nodCurr.Key
        
            If nodCurr = nodFirst.LastSibling Then Exit Do
            Set nodCurr = nodCurr.Next
            If bAbort Then Exit Sub
        Loop
    End If
End Sub

Public Sub WinTrustVerifyNode(sKey$)
    If bAbort Then Exit Sub
    If Not NodeIsValidFile(frmStartupList2.tvwMain.Nodes(sKey)) Then Exit Sub
        
    Dim sFile$, sMD5$, sIcon$
    sFile = frmStartupList2.tvwMain.Nodes(sKey).Text
    If Not FileExists(sFile) Then
        sFile = frmStartupList2.tvwMain.Nodes(sKey).Tag
        If Not FileExists(sFile) Then Exit Sub
    End If
    'Verifying file signature of:
    Status Translate(973) & " " & sFile
    'sMD5 = GetFileMD5(sFile)
    
    Select Case VerifyFileSignature(sFile)
        Case 1: sIcon = "wintrust1"
        Case 0: sIcon = "wintrust3"
        Case -1: Exit Sub
    End Select
    
    frmStartupList2.tvwMain.Nodes(sKey).Image = sIcon
    frmStartupList2.tvwMain.Nodes(sKey).SelectedImage = sIcon
End Sub

Private Function WinVerifyFile(sFile As String) As Boolean
    Dim SignResult As SignResult_TYPE
    SignVerify sFile, 0&, SignResult
    WinVerifyFile = SignResult.isLegit
End Function
