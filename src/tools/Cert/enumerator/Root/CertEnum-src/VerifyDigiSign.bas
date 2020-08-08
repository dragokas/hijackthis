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

' Part of original code for this port was written on C++ by:
' 1) AD13 ( http://forum.sysinternals.com/howto-verify-the-digital-signature-of-a-file_topic19247.html )
' 2) Process Hacker - image verification by wj32 (License: GNU GPL v3) ( http://processhacker.sourceforge.net/doc/verify_8c_source.html )

'Some code examples:
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa382384%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
'https://support.microsoft.com/en-us/kb/323809?wa=wsignin1.0
'https://msdn.microsoft.com/en-us/library/aa382384.aspx
'http://rsdn.ru/forum/src/3152752.hot
'http://rsdn.ru/forum/winapi/2731079.hot
'http://processhacker.sourceforge.net/doc/verify_8c_source.html
'http://eternalwindows.jp/crypto/certverify/certverify03.html

'CERT_CHAIN_POLICY_STATUS structure
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa377188%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396

'Error information:
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa378137(v=vs.85).aspx

' revision 2.2. (17.05.2016)
' Added SHA256 support

#Const UseHashTable = False ' использовать хеш-таблицы? (Maded by Кривоус Анатолий)

Const MAX_PATH As Long = 260&

Enum VB_FILE_ACCESS_MODE
    FOR_READ = 1
    FOR_READ_WRITE = 2
    FOR_OVERWRITE_CREATE = 3
End Enum


Public Type SignResult_TYPE ' данные об ЭЦП
    isSigned     As Boolean ' подписан?
    isLegit      As Boolean ' ЭЦП легитимна?
    isCert       As Boolean ' подписан через каталог безопасности Windows?
    Issuer       As String  ' имя подписанта
    RootCertHash As String
    ShortMessage As String  ' краткое описание результата проверки
    FullMessage  As String  ' полное описание результата проверки
    ReturnCode   As String  ' код возврата WinVerifyTrust
End Type

Public Enum FLAGS_SignVerify
    SV_CheckHoleChain = 1       ' - проверять всю цепочку доверия ( потребуется подключение к сети )
    SV_DoNotUseHashChecking = 2 ' - не использовать проверку по хешу
    SV_DisableCatalogVerify = 4 ' - не использовать проверку по каталогу безопасности
    SV_isDriver = 8             ' - выполнить проверку драйвера на соответствие стандарту WHQL
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
Private Declare Function WinVerifyTrust Lib "Wintrust.dll" (ByVal hwnd As Long, pgActionID As GUID, ByVal pWVTData As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Arguments As Any) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long

Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long


Const WTD_UI_NONE                   As Long = 2&
' проверка отзыва сертификата
Const WTD_REVOKE_NONE               As Long = 0&    ' не проверять сертификаты на предмет отзыва
Const WTD_REVOKE_WHOLECHAIN         As Long = 1&    ' проверять на предмет отзыва сертификаты по цепочке доверия
' способ проверки
Const WTD_CHOICE_CATALOG            As Long = 2&    ' проверка через сравнение с сертификатом, занесенным в локальное хранилище безопасности Windows
Const WTD_CHOICE_FILE               As Long = 1&    ' полная верификация
' флаги
Const WTD_SAFER_FLAG                As Long = 256&   ' ??? (probably, no UI for XP SP2)
Const WTD_REVOCATION_CHECK_NONE     As Long = 16&    ' не проверять цепочку доверия на предмет отзыва сертификата
Const WTD_REVOCATION_CHECK_END_CERT As Long = &H20&  ' проверять на предмет отзыва только конечный сертификат
Const WTD_REVOCATION_CHECK_CHAIN    As Long = &H40&  ' проверять всю цепочку доверия ( требуется подключение к порту 53 TCP/UDP )
Const WTD_REVOCATION_CHECK_CHAIN_EXCLUDE_ROOT As Long = &H80& ' проверять всю цепочку, кроме корневого сертификата
Const WTD_HASH_ONLY_FLAG            As Long = &H200& ' проверка только по хешу
Const WTD_NO_POLICY_USAGE_FLAG      As Long = 4&     ' не брать в рассчет настройки локальной политики безопасности
Const WTD_CACHE_ONLY_URL_RETRIEVAL  As Long = 4096&  ' можно проверять сертификат на предмет отзыва, но только используя данные из локального кеша
Const WTD_LIFETIME_SIGNING_FLAG     As Long = &H800& ' проверять на истечение срока действия сертификата
' действие
Const WTD_STATEACTION_VERIFY        As Long = 1&
Const WTD_STATEACTION_IGNORE        As Long = 0&
Const WTD_STATEACTION_CLOSE         As Long = 2&
' контекст
Const WTD_UICONTEXT_EXECUTE         As Long = 0&
' ошибки
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

' прочее
Const INVALID_HANDLE_VALUE          As Long = -1&
Const ERROR_INSUFFICIENT_BUFFER     As Long = 122&

Dim WINTRUST_ACTION_GENERIC_VERIFY_V2   As GUID
Dim DRIVER_ACTION_VERIFY                As GUID

Public cOut         As Long
Public cErr         As Long
Dim lWow64Old As Long


Public Function SignVerify( _
    FilePath As String, _
    Flags As FLAGS_SignVerify, _
    SignResult As SignResult_TYPE) As Boolean
 
    ' in.  FilePath - путь к PE EXE файлу для проверки
    ' in.  Flags - опции проверки
    ' out. SignResult.ShortMessage - краткое описание статуса проверки
    ' out. SignResult.FullMessage - послное описание результата проверки
    ' out. SignResult.ReturnCode - код, возвращаемый функцией WinVerifyTrust ( по нему смотреть статус проверки )
    
    ' RETURN - Вернет true, если целостность исполняемого файла подтверждена, невзирая на:
    ' - возможные запреты в настройках локальной политики
    ' - самоподписанный тип сертификата ( если не указана опция CheckHoleChain = true и данные об отзыве не были закешированы )
    ' - проверка на просрочку действия сертификата не выполняется. Если нужна, добавляем флаг WTD_LIFETIME_SIGNING_FLAG
    
    ' P.S. Для запрета самоподписанных сертификатов, удалите флаг CERT_E_UNTRUSTEDROOT
    ' Для еще менее строгой проверки ( запрете чтения из кеша данных об отзыве ), замените флаг WTD_CACHE_ONLY_URL_RETRIEVAL на WTD_REVOCATION_CHECK_NONE.
    ' Обратите внимание, что отзыв сертификата - это исключительная процедура, которая проводится например, когда ЭЦП была украдена или использована во вредоносном ПО.
    
    ' in. Flags (можно комбинировать друг с другом через OR):
    ' SV_CheckHoleChain = 1       ' - проверять всю цепочку доверия ( потребуется подключение к сети )
    ' SV_DoNotUseHashChecking = 2 ' - не использовать проверку по хешу
    ' SV_DisableCatalogVerify = 4 ' - не использовать проверку по каталогу безопасности
    ' SV_isDriver = 8             ' - выполнить проверку драйвера на соответствие стандарту WHQL
    
    Dim InfoStruct               As CATALOG_INFO
    Dim WintrustStructure        As WINTRUST_DATA
    Dim WintrustCatalogStructure As WINTRUST_CATALOG_INFO
    Dim WintrustFileStructure    As WINTRUST_FILE_INFO
    
    Static MajorMinor       As Single
    Static IsVistaAndLater  As Boolean
    Static IsWin8AndLater   As Boolean
    Static SignCache()      As SignResult_TYPE
    Static SC_pos           As Long
    #If UseHashTable Then
        'Static oSignIndex As clsTrickHashTable
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
    
    If 0 = ObjPtr(oSignIndex) Then
        #If UseHashTable Then
            'Set oSignIndex = New clsTrickHashTable
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
    
    With SignResult     'обнуление результата проверки
        .ReturnCode = 0
        .FullMessage = vbNullString
        .ShortMessage = vbNullString
        '.Signer = vbNullString
        .isSigned = False
        .isLegit = False
    End With
    
    ToggleWow64FSRedirection True
    
    SC_pos = SC_pos + 1
    If UBound(SignCache) < SC_pos Then ReDim Preserve SignCache(UBound(SignCache) + 100)
    oSignIndex.Add FilePath, SC_pos
    
    If MajorMinor = 0 Then  'not cached
        hLib = LoadLibrary(StrPtr("Wintrust.dll"))
        If hLib = 0 Then
            WriteCon "NOT SUPPORTED. LastDllErr=0x" & Hex(err.LastDllError)
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

    If (Flags And SV_isDriver) Then
        ActionGuid = DRIVER_ACTION_VERIFY
    Else
        ActionGuid = WINTRUST_ACTION_GENERIC_VERIFY_V2
    End If
    
    InfoStruct.cbStruct = Len(InfoStruct)
    WintrustFileStructure.cbStruct = Len(WintrustFileStructure)
    
    If MajorMinor >= 6.2 Then
        ' Получаем контекст для процедуры проверки подписи
        CryptCATAdminAcquireContext2 Context, VarPtr(DRIVER_ACTION_VERIFY), StrPtr(BCRYPT_SHA256_ALGORITHM), 0&, 0&
    End If
    
    If Context = 0 Then
        If Not (CBool(CryptCATAdminAcquireContext(Context, VarPtr(ActionGuid), 0&))) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminAcquireContext", FilePath
            SignCache(SC_pos) = SignResult
            Exit Function
        End If
    End If
    
    ' Открываем файл
    ToggleWow64FSRedirection False, FilePath
    
    OpenW FilePath, FOR_READ, FileHandle
    
    ToggleWow64FSRedirection True
    
    If (INVALID_HANDLE_VALUE = FileHandle) Then
        WriteError err, SignResult.ShortMessage, "CreateFile"
        CryptCATAdminReleaseContext Context, 0&
        SignCache(SC_pos) = SignResult
        Exit Function
    End If
    

    If MajorMinor >= 6.2 Then
        ' Получаем размер, необходимый для хеша
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
        ' Получаем размер, необходимый для хеша
        CryptCATAdminCalcHashFromFileHandle FileHandle, HashSize, ByVal 0&, 0&
    
        If (HashSize = 0&) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminCalcHashFromFileHandle", FilePath
            CryptCATAdminReleaseContext Context, 0&
            SignCache(SC_pos) = SignResult
            CloseHandle FileHandle: FileHandle = 0
            Exit Function
        End If

        ' Выделяем память
        ReDim aBuf(HashSize - 1&)

        ' Подсчитываем хеш
        If Not CBool(CryptCATAdminCalcHashFromFileHandle(FileHandle, HashSize, aBuf(0), 0&)) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminCalcHashFromFileHandle", FilePath
            CryptCATAdminReleaseContext Context, 0&
            SignCache(SC_pos) = SignResult
            CloseHandle FileHandle: FileHandle = 0
            Exit Function
        End If
    End If
    
    
    ' Преобразовываем хеш в строку
    For i = 0& To UBound(aBuf)
        MemberTag = MemberTag & Right$("0" & Hex(aBuf(i)), 2&)
    Next
    'Debug.Print MemberTag

    'here is M$ bug with  C:\WINDOWS\system32\catroot2\{GUID} file. // Need to avoid.

    ' Получаем каталог для нашего контекста
    If Not HasVulnerability() Then
        CatalogContext = CryptCATAdminEnumCatalogFromHash(Context, aBuf(0), HashSize, 0&, ByVal 0&)
        Debug.Print "CatalogContext: " & CatalogContext
        Stop
    End If

    If (CatalogContext) Then
        ' Если не можем получить информацию
        If Not CBool(CryptCATCatalogInfoFromContext(CatalogContext, InfoStruct, 0&)) Then
            WriteError err, SignResult.ShortMessage, "CryptCATCatalogInfoFromContext"
            ' Освобождаем контекст
            CryptCATAdminReleaseCatalogContext Context, CatalogContext, 0&
            SignCache(SC_pos) = SignResult
            CatalogContext = 0&
        End If
        
        'sBuf = String$(MAX_PATH, vbNullChar)
        'memcpy ByVal StrPtr(sBuf), InfoStruct.wszCatalogFile(0), MAX_PATH
        'sBuf = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
        'Debug.Print "Hash: " & sBuf
    End If
    
    ' Если получили валидный контекст, проверяем подпись через каталог.
    ' Иначе, пытаемся проверить внутреннюю подпись файла:
    If (CatalogContext = 0& Or (Flags And SV_DisableCatalogVerify)) Then
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
            .fdwRevocationChecks = IIf(Flags And SV_CheckHoleChain, WTD_REVOKE_WHOLECHAIN, WTD_REVOKE_NONE)   ' проверка на отзыв сертификата
            If Flags And SV_CheckHoleChain Then
                .dwProvFlags = .dwProvFlags Or WTD_REVOCATION_CHECK_CHAIN
            Else
                ' брать данные о проверке цепочки сертификатов только из локального кеша, если сохранены ( >= Vista ). Не использовать подключение к сети.
                .dwProvFlags = .dwProvFlags Or IIf(IsVistaAndLater, WTD_CACHE_ONLY_URL_RETRIEVAL, WTD_REVOCATION_CHECK_NONE)
            End If
            '.dwProvFlags = .dwProvFlags Or WTD_NO_POLICY_USAGE_FLAG                                          ' не проверять назначение сертификата (отключено)
            .dwProvFlags = .dwProvFlags Or WTD_LIFETIME_SIGNING_FLAG                                          ' признавать недействительными просроченные подписи
            If 0 = (Flags And SV_DoNotUseHashChecking) Then .dwProvFlags = .dwProvFlags Or WTD_HASH_ONLY_FLAG ' проверка только по хешу
            If IsVistaAndLater Then .dwProvFlags = .dwProvFlags Or WTD_SAFER_FLAG                             ' без UI (XP SP2)
        End With
    
    Else
        ' Получили информацию из каталога. Проверяем валидность через него.
        SignResult.isSigned = True
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
            .dwProvFlags = 0&
            .dwUIContext = WTD_UICONTEXT_EXECUTE
        End With
        
        ' Заполняем структуру каталога
        With WintrustCatalogStructure               'WINTRUST_CATALOG_INFO
            .cbStruct = Len(WintrustCatalogStructure)
            .dwCatalogVersion = 0&
            .pcwszCatalogFilePath = VarPtr(InfoStruct.wszCatalogFile(0))
            .pcwszMemberTag = StrPtr(MemberTag)
            .pcwszMemberFilePath = StrPtr(FilePath)
            .hMemberFile = 0&
        End With
    End If
    
    'ToggleWow64FSRedirection False, FilePath
    
    ' Вызываем функцию проверки подписи
    ReturnVal = WinVerifyTrust(INVALID_HANDLE_VALUE, ActionGuid, VarPtr(WintrustStructure)) ' INVALID_HANDLE_VALUE означает неинтерактивную проверку (без UI)

    'ToggleWow64FSRedirection True

    'Debug.Print WintrustStructure.hWVTStateData    ' -> идентификатор, по которому можно извлечь отдельные подписи
    
    GetSignerInfo WintrustStructure.hWVTStateData
    
    'Stop
    
    ' разрешить самоподписанный сертификат ( CERT_E_UNTRUSTEDROOT - игнорировать отсутствие недоверия к конечному сертификату )
    'ReturnFlag = ((ReturnVal = 0) Or (ReturnVal = CRYPT_E_SECURITY_SETTINGS) Or (ReturnVal = CERT_E_UNTRUSTEDROOT))
    ReturnFlag = (ReturnVal = 0)
    
    With SignResult
    
        If True = ReturnFlag Then
            .isSigned = True
        Else
            ' проверка наличия подписи
            If (CatalogContext = 0& Or (Flags And SV_DisableCatalogVerify)) Then
                '.isSigned = IsSignPresent(FilePath)
            End If
        End If

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
    
        ' Прочие коды ошибок можно посмотреть на MSDN:
        ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa377188%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
        ' Это не исчерпывающий список.
    
        .FullMessage = ErrMessageText(ReturnVal)
    
        .ReturnCode = ReturnVal
        .isLegit = ReturnFlag
        SignVerify = ReturnFlag
    
        WriteCon .ShortMessage
        WriteCon "FormatMessage: " & .FullMessage
    
    End With
    
    ' Освобождаем контекст
    If (CatalogContext) Then
        CryptCATAdminReleaseCatalogContext Context, CatalogContext, 0&
    End If
    
    ' Очистка памяти, использованная провайдером во время проверки подписи
    WintrustStructure.dwStateAction = WTD_STATEACTION_CLOSE
    WinVerifyTrust INVALID_HANDLE_VALUE, ActionGuid, VarPtr(WintrustStructure)
    
    ' Закрытие файла и освобождение контекста
    CloseHandle FileHandle
    CryptCATAdminReleaseContext Context, 0&
    
    SignCache(SC_pos) = SignResult
    
    ToggleWow64FSRedirection True
End Function




Public Function SignVerify_2( _
    FilePath As String, _
    Flags As FLAGS_SignVerify, _
    SignResult As SignResult_TYPE) As Boolean

    On Error GoTo ErrorHandler

    'AppendErrorLogCustom "VerifyDigiSign.SignVerify - Begin"

    ' in.  FilePath - путь к PE EXE файлу для проверки
    ' in.  Flags - опции проверки
    ' out. SignResult.ShortMessage - краткое описание статуса проверки
    ' out. SignResult.FullMessage - послное описание результата проверки
    ' out. SignResult.ReturnCode - код, возвращаемый функцией WinVerifyTrust ( по нему смотреть статус проверки )

    ' RETURN - Вернет true, если целостность исполняемого файла подтверждена, невзирая на:
    ' - возможные запреты в настройках локальной политики
    ' - самоподписанный тип сертификата ( если не указана опция CheckHoleChain = true и данные об отзыве не были закешированы )
    ' - проверка на просрочку действия сертификата не выполняется. Если нужна, добавляем флаг WTD_LIFETIME_SIGNING_FLAG

    ' P.S. Для запрета самоподписанных сертификатов, удалите флаг CERT_E_UNTRUSTEDROOT
    ' Для еще менее строгой проверки ( запрете чтения из кеша данных об отзыве ), замените флаг WTD_CACHE_ONLY_URL_RETRIEVAL на WTD_REVOCATION_CHECK_NONE.
    ' Обратите внимание, что отзыв сертификата - это исключительная процедура, которая проводится например, когда ЭЦП была украдена или использована во вредоносном ПО.

    ' in. Flags (можно комбинировать друг с другом через OR):
    ' SV_CheckHoleChain = 1       ' - проверять всю цепочку доверия ( потребуется подключение к сети )
    ' SV_DoNotUseHashChecking = 2 ' - не использовать проверку по хешу
    ' SV_DisableCatalogVerify = 4 ' - не использовать проверку по каталогу безопасности
    ' SV_isDriver = 8             ' - выполнить проверку драйвера на соответствие стандарту WHQL

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
        'Static oSignIndex As clsTrickHashTable
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

    With SignResult     'обнуление результата проверки
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
            'WriteCon "NOT SUPPORTED. LastDllErr=0x" & Hex(err.LastDllError)
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

    If (Flags And SV_isDriver) Then
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
        ' Получаем контекст для процедуры проверки подписи
        CryptCATAdminAcquireContext2 Context, VarPtr(DRIVER_ACTION_VERIFY), StrPtr(BCRYPT_SHA256_ALGORITHM), 0&, 0&
    End If

    If Context = 0 Then
        If Not (CBool(CryptCATAdminAcquireContext(Context, VarPtr(DRIVER_ACTION_VERIFY), 0&))) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminAcquireContext", FilePath
            SignCache(SC_pos) = SignResult
            Exit Function
        End If
    End If

'    If Not FileExists(FilePath) Then
'        CryptCATAdminReleaseContext Context, 0&
'        SignCache(SC_pos) = SignResult
'        Exit Function
'    End If

    ' Открываем файл
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

    ' Проверяем наличие внутренней подписи
    'SignResult.isSigned = IsSignPresent(FilePath) ', FileHandle)

    If MajorMinor >= 6.2 Then
        ' Получаем размер, необходимый для хеша
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
        ' Получаем размер, необходимый для хеша
        CryptCATAdminCalcHashFromFileHandle FileHandle, HashSize, ByVal 0&, 0&

        If (HashSize = 0&) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminCalcHashFromFileHandle", FilePath
            CryptCATAdminReleaseContext Context, 0&
            SignCache(SC_pos) = SignResult
            CloseHandle FileHandle: FileHandle = 0
            Exit Function
        End If

        ' Выделяем память
        ReDim aBuf(HashSize - 1&)

        ' Подсчитываем хеш
        If Not CBool(CryptCATAdminCalcHashFromFileHandle(FileHandle, HashSize, aBuf(0), 0&)) Then
            WriteError err, SignResult.ShortMessage, "CryptCATAdminCalcHashFromFileHandle", FilePath
            CryptCATAdminReleaseContext Context, 0&
            SignCache(SC_pos) = SignResult
            CloseHandle FileHandle: FileHandle = 0
            Exit Function
        End If
    End If

    ' Преобразовываем хеш в строку
    For i = 0& To UBound(aBuf)
        MemberTag = MemberTag & Right$("0" & Hex(aBuf(i)), 2&)
    Next
    'Debug.Print MemberTag

    'here is M$ bug with  C:\WINDOWS\system32\catroot2\{GUID} file. // Need to avoid.

    ' Получаем каталог для нашего контекста
    If Not HasVulnerability() Then
        CatalogContext = CryptCATAdminEnumCatalogFromHash(Context, aBuf(0), HashSize, 0&, ByVal 0&)
        Debug.Print "CatalogContext: " & CatalogContext
    End If

    If (CatalogContext) Then
        ' Если не можем получить информацию
        If Not CBool(CryptCATCatalogInfoFromContext(CatalogContext, InfoStruct, 0&)) Then
            WriteError err, SignResult.ShortMessage, "CryptCATCatalogInfoFromContext", FilePath
            ' Освобождаем контекст
            CryptCATAdminReleaseCatalogContext Context, CatalogContext, 0&
            SignCache(SC_pos) = SignResult
            CatalogContext = 0&
        End If

        'sBuf = String$(MAX_PATH, vbNullChar)
        'memcpy ByVal StrPtr(sBuf), InfoStruct.wszCatalogFile(0), MAX_PATH
        'sBuf = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
        'Debug.Print "Hash: " & sBuf
    End If

    ' Если получили валидный контекст, проверяем подпись через каталог.
    ' Иначе (если есть Embedded signature или установлен флаг "Игнорировать проверку через каталог"), пытаемся проверить внутреннюю подпись файла:
    If (CatalogContext = 0& Or (Flags And SV_DisableCatalogVerify)) Or SignResult.isSigned Then
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
            .fdwRevocationChecks = IIf(Flags And SV_CheckHoleChain, WTD_REVOKE_WHOLECHAIN, WTD_REVOKE_NONE)   ' проверка на отзыв сертификата
            If Flags And SV_CheckHoleChain Then
                .dwProvFlags = .dwProvFlags Or WTD_REVOCATION_CHECK_CHAIN
            Else
                ' брать данные о проверке цепочки сертификатов только из локального кеша, если сохранены ( >= Vista ). Не использовать подключение к сети.
                .dwProvFlags = .dwProvFlags Or IIf(IsVistaAndLater, WTD_CACHE_ONLY_URL_RETRIEVAL, WTD_REVOCATION_CHECK_NONE)
            End If
            '.dwProvFlags = .dwProvFlags Or WTD_NO_POLICY_USAGE_FLAG                                          ' не проверять назначение сертификата (отключено)
            .dwProvFlags = .dwProvFlags Or WTD_LIFETIME_SIGNING_FLAG                                          ' признавать недействительными просроченные подписи
            If 0 = (Flags And SV_DoNotUseHashChecking) Then .dwProvFlags = .dwProvFlags Or WTD_HASH_ONLY_FLAG ' проверка только по хешу
            If IsVistaAndLater Then .dwProvFlags = .dwProvFlags Or WTD_SAFER_FLAG                             ' без UI (XP SP2)
        End With

    Else
        ' Получили информацию из каталога. Проверяем валидность через него.
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
            .dwProvFlags = 0&
            .dwUIContext = WTD_UICONTEXT_EXECUTE
        End With

        ' Заполняем структуру каталога
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

    'If FilePath = "C:\Windows\system32\wbem\WmiApSrv.exe" Then Stop

    ' Вызываем функцию проверки подписи
    ReturnVal = WinVerifyTrust(INVALID_HANDLE_VALUE, ActionGuid, VarPtr(WintrustStructure)) ' INVALID_HANDLE_VALUE означает неинтерактивную проверку (без UI)
    'Stop
    ToggleWow64FSRedirection True

    'Debug.Print WintrustStructure.hWVTStateData    ' -> идентификатор, по которому можно извлечь отдельные подписи

    GetSignerInfo WintrustStructure.hWVTStateData ', SignResult.Issuer, SignResult.RootCertHash

    ' разрешить самоподписанный сертификат ( CERT_E_UNTRUSTEDROOT - игнорировать отсутствие недоверия к конечному сертификату )
    'ReturnFlag = ((ReturnVal = 0) Or (ReturnVal = CRYPT_E_SECURITY_SETTINGS) Or (ReturnVal = CERT_E_UNTRUSTEDROOT))
    ReturnFlag = (ReturnVal = 0)

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

        ' Прочие коды ошибок можно посмотреть на MSDN:
        ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa377188%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
        ' https://msdn.microsoft.com/en-us/library/ee488436.aspx
        ' Это не исчерпывающий список.

        .FullMessage = ErrMessageText(ReturnVal)

        If .FullMessage <> "Операция успешно завершена." And .FullMessage <> "В этом объекте нет подписи." Then
            Debug.Print FilePath & vbTab & vbTab & .FullMessage
        End If

        .ReturnCode = ReturnVal
        .isLegit = ReturnFlag
        'SignVerify = ReturnFlag

        'WriteCon .ShortMessage
        'WriteCon "FormatMessage: " & .FullMessage

    End With

    ' Освобождаем контекст
    If (CatalogContext) Then
        CryptCATAdminReleaseCatalogContext Context, CatalogContext, 0&
    End If

    ' Очистка памяти, использованная провайдером во время проверки подписи
    WintrustStructure.dwStateAction = WTD_STATEACTION_CLOSE
    WinVerifyTrust INVALID_HANDLE_VALUE, ActionGuid, VarPtr(WintrustStructure)

    ' Закрытие файла и освобождение контекста
    CloseHandle FileHandle: FileHandle = 0
    CryptCATAdminReleaseContext Context, 0&

    SignCache(SC_pos) = SignResult

    'AppendErrorLogCustom "VerifyDigiSign.SignVerify - End"
    Exit Function

ErrorHandler:
    'AppendErrorLogFormat Now, err, "VerifyDigiSign.SignVerify", "File:", FilePath
    Debug.Print err, "SignVerify", FilePath
    ToggleWow64FSRedirection True
End Function

Function HasVulnerability() As Boolean
    On Error GoTo ErrHandler
    Static isInit       As Boolean
    Static VulnStatus   As Boolean
    
    If isInit Then
        HasVulnerability = VulnStatus
        Exit Function
    Else
        isInit = True
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
                VulnStatus = True: HasVulnerability = True: Exit Function
            End If
        End If
        sFile = Dir$()
    Loop
    Exit Function
ErrHandler:
    Debug.Print err, "HasVulnerability"
    
End Function

Public Function ErrMessageText(lCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
    Const FORMAT_MESSAGE_FROM_HMODULE   As Long = &H800&
    
    Dim sRtrnMessage   As String
    Dim lret           As Long
    'Dim hLib           As Long
    
    sRtrnMessage = String$(MAX_PATH, vbNullChar)
    lret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lCode, 0&, StrPtr(sRtrnMessage), MAX_PATH, ByVal 0&)
    If lret > 0 Then
        ErrMessageText = Left$(sRtrnMessage, lret)
        ErrMessageText = Replace$(ErrMessageText, vbCrLf, vbNullString)
    End If
End Function

' Запишет ошибку с описанием в консоль и в ShortMessage (RetErrMessage)
Function WriteError(err As ErrObject, RetErrMessage As String, FunctionName As String, Optional FileName As String)
    Dim ErrNumber As Long
    ErrNumber = err.LastDllError
    
    If &H800700C1 = ErrNumber Then
        ' если получили "%1 не является приложением Win32.", и при этом в PE EXE есть указатель на структуру SecurityDir,
        ' значит ЭЦП была повреждена
        ' https://chentiangemalc.wordpress.com/2014/08/01/case-of-the-server-returned-a-referral/

        'If IsSignPresent(FileName) Then
        '    RetErrMessage = "Signature is damaged!"
        'Else
            RetErrMessage = ErrMessageText(ErrNumber)
        'End If
    Else
        RetErrMessage = ErrMessageText(ErrNumber)
    End If
    
    WriteCon "Error in " & FunctionName & ": " & "0x" & Hex(ErrNumber) & ". " & RetErrMessage & ". File: " & FileName
    'AppendErrorLogFormat Now, err, "VerifyDigiSign." & FunctionName & ": " & "0x" & Hex(ErrNumber) & ". " & RetErrMessage & " File: " & FileName
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

Public Function ToggleWow64FSRedirection(bEnable As Boolean, Optional PathNecessity As String, Optional OldStatus As Boolean) As Boolean
    'Static lWow64Old        As Long    'Warning: do not use initialized variables for this API !
                                        'Static variables is not allowed !
                                        'lWow64Old is now declared globally
    'True - enable redirector
    'False - disable redirector

    'OldStatus: current state of redirection
    'True - redirector was enabled
    'False - redirector was disabled

    'Return value is:
    'true if success

    Static IsNotRedirected  As Boolean
    Dim lr                  As Long
    Dim sWinDir$

    OldStatus = Not IsNotRedirected

    If Not IsWOW64 Then Exit Function

    If Len(PathNecessity) <> 0 Then
        sWinDir = Environ("WinDir")
        If StrComp(Left$(PathNecessity, Len(sWinDir)), sWinDir, vbTextCompare) <> 0 Then Exit Function
    End If

    If bEnable Then
        If IsNotRedirected Then
            lr = Wow64RevertWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = False
        End If
    Else
        If Not IsNotRedirected Then
            lr = Wow64DisableWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = True
        End If
    End If
End Function


Public Function OpenW(FileName As String, Access As VB_FILE_ACCESS_MODE, retHandle As Long, Optional MountToMemory As Boolean) As Boolean '// TODO: MountToMemory
    Const OPEN_EXISTING     As Long = 3&
    Const CREATE_ALWAYS     As Long = 2&
    Const GENERIC_READ      As Long = &H80000000
    Const GENERIC_WRITE     As Long = &H40000000
    Const FILE_SHARE_READ   As Long = 1&
    Const FILE_SHARE_WRITE  As Long = 2&
    Const FILE_SHARE_DELETE As Long = 4&
    
    If Access = FOR_READ Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_OVERWRITE_CREATE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_READ_WRITE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        WriteCon "Wrong access mode!", cErr
    End If
    OpenW = (INVALID_HANDLE_VALUE <> retHandle)
End Function

                                                                  'do not change Variant type at all or you will die ^_^
Public Function GetW(hFile As Long, pos As Long, Optional vOut As Variant, Optional vOutPtr As Long, Optional cbToRead As Long) As Long
                                                                  
    On Error GoTo ErrorHandler
    
    Const NO_ERROR                  As Long = 0&
    Const FILE_BEGIN                As Long = 0&
    Const FILE_CURRENT              As Long = 1&
    Const FILE_END                  As Long = 2&
    Const INVALID_SET_FILE_POINTER  As Long = &HFFFFFFFF
    
    Dim lBytesRead  As Long
    Dim lr          As Long
    Dim ptr         As Long
    Dim vType       As Long
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If INVALID_SET_FILE_POINTER <> SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then
        If NO_ERROR = err.LastDllError Then
            vType = VarType(vOut)
            If 0 <> cbToRead Then   'vbError = vType
                lr = ReadFile(hFile, vOutPtr, cbToRead, lBytesRead, 0&)
            ElseIf vbString = vType Then
                lr = ReadFile(hFile, StrPtr(vOut), Len(vOut), lBytesRead, 0&)
                vOut = StrConv(vOut, vbUnicode)
                If Len(vOut) <> 0 Then vOut = Left$(vOut, Len(vOut) \ 2)
            Else
                'do a bit of magik :)
                memcpy ptr, ByVal VarPtr(vOut) + 8, 4& 'VT_BYREF
                Select Case vType
                Case vbLong
                    lr = ReadFile(hFile, ptr, 4&, lBytesRead, 0&)
                Case vbInteger
                    lr = ReadFile(hFile, ptr, 2&, lBytesRead, 0&)
                Case vbCurrency
                    lr = ReadFile(hFile, ptr, 8&, lBytesRead, 0&)
                Case Else
                    WriteCon "Error! GetW for type #" & VarType(vOut) & " of buffer is not supported.", cErr
                End Select
            End If
            If 0 = lr Then
                WriteCon "Cannot read file!", cErr: err.Raise 52
            Else
                GetW = True
            End If
        Else
            WriteCon "Cannot set file pointer!", cErr: err.Raise 52
        End If
    Else
        WriteCon "Cannot set file pointer!", cErr: err.Raise 52
    End If
    
    Exit Function
ErrorHandler:
    WriteCon "Error #" & err.Number & ". LastDll=" & err.LastDllError & ". " & err.Description, cErr
    'ExitProcess 1
End Function

Public Function LOFW(hFile As Long) As Currency
    On Error Resume Next
    Dim lr          As Long
    Dim FileSize    As Currency
    
    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then LOFW = FileSize * 10000
        End If
    End If
End Function

Function IsWOW64() As Boolean
    Dim hModule As Long, procAddr As Long, lIsWin64 As Long
    
    hModule = LoadLibrary(StrPtr("kernel32.dll"))
    If hModule Then
        procAddr = GetProcAddress(hModule, "IsWow64Process")
        If procAddr <> 0 Then
            IsWow64Process GetCurrentProcess(), lIsWin64
            IsWOW64 = CBool(lIsWin64)
        End If
        FreeLibrary hModule
    End If
End Function
