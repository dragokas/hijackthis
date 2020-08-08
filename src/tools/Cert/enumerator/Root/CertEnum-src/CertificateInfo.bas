Attribute VB_Name = "SignerInfo"
Option Explicit

'https://support.microsoft.com/en-us/kb/323809
'https://www.sysadmins.lv/blog-ru/certificate-trust-list-ctl-v-powershell.aspx
'https://www.sysadmins.lv/blog-ru/chto-v-oide-tebe-moem.aspx
'certmgr.msc
'certmgr.exe (Windows SDK)


Type CRYPT_KEY_PROV_INFO
    pwszContainerName As Long
    
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type CRYPTOAPI_BLOB
    cbData As Long
    pbData As Long ' ptr -> BYTE[]
End Type
'Alias for:
'CRYPT_INTEGER_BLOB, *PCRYPT_INTEGER_BLOB, CRYPT_UINT_BLOB, *PCRYPT_UINT_BLOB, CRYPT_OBJID_BLOB, *PCRYPT_OBJID_BLOB, CERT_NAME_BLOB,
'CERT_RDN_VALUE_BLOB, *PCERT_NAME_BLOB, *PCERT_RDN_VALUE_BLOB, CERT_BLOB, *PCERT_BLOB, CRL_BLOB, *PCRL_BLOB, DATA_BLOB, *PDATA_BLOB,
'CRYPT_DATA_BLOB, *PCRYPT_DATA_BLOB, CRYPT_HASH_BLOB, *PCRYPT_HASH_BLOB, CRYPT_DIGEST_BLOB, *PCRYPT_DIGEST_BLOB, CRYPT_DER_BLOB,
'PCRYPT_DER_BLOB, CRYPT_ATTR_BLOB, *PCRYPT_ATTR_BLOB;

Type CRYPT_BIT_BLOB
    cbData      As Long
    pbData      As Long ' ptr -> BYTE[]
    cUnusedBits As Long
End Type

Type CRYPT_ALGORITHM_IDENTIFIER
    pszObjId    As Long ' ptr -> STR
    Parameters  As CRYPTOAPI_BLOB ' CRYPT_OBJID_BLOB
End Type

Type CERT_PUBLIC_KEY_INFO
    Algorithm   As CRYPT_ALGORITHM_IDENTIFIER
    PublicKey   As CRYPT_BIT_BLOB
End Type

Type CERT_INFO
    dwVersion               As Long
    SerialNumber            As CRYPTOAPI_BLOB
    SignatureAlgorithm      As CRYPT_ALGORITHM_IDENTIFIER
    Issuer                  As CRYPTOAPI_BLOB
    NotBefore               As FILETIME
    NotAfter                As FILETIME
    Subject                 As CRYPTOAPI_BLOB ' CERT_NAME_BLOB
    SubjectPublicKeyInfo    As CERT_PUBLIC_KEY_INFO
    IssuerUniqueId          As CRYPT_BIT_BLOB
    SubjectUniqueId         As CRYPT_BIT_BLOB
    cExtension              As Long
    rgExtension             As Long ' ptr -> CERT_EXTENSION
End Type

Type CERT_CONTEXT
    dwCertEncodingType      As Long
    pbCertEncoded           As Long ' ptr -> encoded certificate
    cbCertEncoded           As Long
    pCertInfo               As Long ' ptr -> PCERT_INFO
    hCertStore              As Long
End Type

Type CRYPT_PROVIDER_CERT
    cbStruct                As Long
    pCert                   As Long ' ptr -> CERT_CONTEXT
    fCommercial             As Long
    fTrustedRoot            As Long
    fSelfSigned             As Long
    fTestCert               As Long
    dwRevokedReason         As Long
    dwConfidence            As Long
    dwError                 As Long
    pTrustListContext       As Long ' ptr -> CTL_CONTEXT
    fTrustListSignerCert    As Long
    pCtlContext             As Long ' ptr -> CTL_CONTEXT
    dwCtlError              As Long
    fIsCyclic               As Long
    pChainElement           As Long ' ptr -> CERT_CHAIN_ELEMENT
End Type

Type CRYPT_PROVIDER_SGNR
    cbStruct                As Long
    sftVerifyAsOf           As FILETIME
    csCertChain             As Long
    pasCertChain            As Long ' ptr -> CRYPT_PROVIDER_CERT
    dwSignerType            As Long
    psSigner                As Long ' ptr -> CMSG_SIGNER_INFO
    dwError                 As Long
    csCounterSigners        As Long
    pasCounterSigners       As Long ' ptr -> array of CRYPT_PROVIDER_SGNR
    pChainContext           As Long ' ptr -> CERT_CHAIN_CONTEXT
End Type

Type CRYPT_PROVIDER_DATA
    cbStruct                As Long
    pWintrustData           As Long ' ptr -> WINTRUST_DATA
    fOpenedFile             As Long ' BOOL
    hWndParent              As Long
    pgActionID              As Long
    hprov                   As Long ' HCRYPTPROV
    dwError                 As Long
    dwRegSecuritySettings   As Long
    dwRegPolicySettings     As Long
    psPfns                  As Long ' ptr -> CRYPT_PROVIDER_FUNCTIONS
    cdwTrustStepErrors      As Long
    padwTrustStepErrors     As Long ' ptr
    chStores                As Long
    pahStores               As Long ' ptr -> HCERTSTORE
    dwEncoding              As Long
    hMsg                    As Long ' HCRYPTMSG
    csSigners               As Long
    pasSigners              As Long ' ptr -> CRYPT_PROVIDER_SGNR
    csProvPrivData          As Long
    pasProvPrivData         As Long ' ptr -> array of CRYPT_PROVIDER_PRIVDATA structures
    dwSubjectChoice         As Long
    pPDSip                  As Long ' ptr -> PROVDATA_SIP
    pszUsageOID             As Long ' ptr -> char
    fRecallWithState        As Long ' BOOL
    sftSystemTime           As FILETIME
    pszCTLSignerUsageOID    As Long ' ptr -> char
    dwProvFlags             As Long
    dwFinalError            As Long
    pRequestUsage           As Long ' ptr -> PCERT_USAGE_MATCH
    dwTrustPubSettings      As Long
    dwUIStateFlags          As Long
    pUnknown1               As Long 'undocumented (Win 7+)
    pUnknown2               As Long 'undocumented (Win 7+)
End Type

Private Declare Function WTHelperProvDataFromStateData Lib "Wintrust.dll" (ByVal StateData As Long) As Long 'CRYPT_PROVIDER_DATA
Private Declare Function WTHelperGetProvSignerFromChain Lib "Wintrust.dll" (ByVal pProvData As Long, ByVal idxSigner As Long, ByVal fCounterSigner As Long, ByVal idxCounterSigner As Long) As Long
Private Declare Function CertDuplicateCertificateContext Lib "Crypt32.dll" (ByVal pCertContext As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "Crypt32.dll" (ByVal pCertContext As Long) As Long
Private Declare Function CertNameToStr Lib "Crypt32.dll" Alias "CertNameToStrW" (ByVal dwCertEncodingType As Long, ByVal pName As Long, ByVal dwStrType As Long, ByVal psz As Long, ByVal csz As Long) As Long
Private Declare Function CertGetCertificateContextProperty Lib "Crypt32.dll" (ByVal pCertContext As Long, ByVal dwPropId As Long, pvData As Any, pcbData As Long) As Long
Private Declare Function CertGetNameString Lib "Crypt32.dll" Alias "CertGetNameStringW" (ByVal pCertContext As Long, ByVal dwType As Long, ByVal dwFlags As Long, pvTypePara As Any, ByVal pszNameString As Long, ByVal cchNameString As Long) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long) As Long
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (arr() As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (pSrc As Any, pDst As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Const X509_ASN_ENCODING As Long = 1&
Private Const CERT_X500_NAME_STR As Long = 3&

Private Const CERT_HASH_PROP_ID As Long = 3&
Private Const CERT_SIGNATURE_HASH_PROP_ID As Long = 15&
Private Const CERT_SUBJECT_PUBLIC_KEY_MD5_HASH_PROP_ID As Long = 25&

Private Const CERT_NAME_SIMPLE_DISPLAY_TYPE As Long = 4&
Private Const CERT_NAME_ISSUER_FLAG As Long = 1&



Public Function GetSignerInfo(StateData As Long) As Boolean
    
    Dim NumberOfSignatures As Long
    Dim Signatures() As Long
    Dim i As Long
    
    GetSignaturesFromStateData StateData, Signatures, NumberOfSignatures
    
    'CERT_HASH_PROP_ID
    'CERT_SHA1_HASH_PROP_ID
    'CERT_SIGNATURE_HASH_PROP_ID
    'CERT_SUBJECT_PUBLIC_KEY_MD5_HASH_PROP_ID   ' certificate's public key.
    
    If NumberOfSignatures Then
        For i = 0 To NumberOfSignatures - 1
            GetSignerNameFromCertificate Signatures(i)
            Debug.Print ExtractPropertyFromCertificate(Signatures(i), CERT_HASH_PROP_ID)
            Debug.Print "Issuer = " & GetCertString(Signatures(i), CERT_NAME_SIMPLE_DISPLAY_TYPE, CERT_NAME_ISSUER_FLAG)
        Next
    End If
    
    For i = 0 To NumberOfSignatures - 1
        CertFreeCertificateContext Signatures(i)
    Next
    
End Function

Function WINTRUST_Free(ptr As Long) As Long
    If 0 <> ptr Then HeapFree GetProcessHeap(), 0, ptr
End Function

Sub GetSignaturesFromStateData(StateData As Long, Signatures() As Long, NumberOfSignatures As Long)
    'Signatures() return pointers to CERT_CONTEXT

    Dim pProvData   As Long
    Dim ProvData    As CRYPT_PROVIDER_DATA
    Dim pCPSigner   As Long
    Dim CPSigner    As CRYPT_PROVIDER_SGNR
    Dim CPCERT()    As CRYPT_PROVIDER_CERT
    Dim lpOldPt     As Long
    Dim lpSA        As Long
    Dim i As Long
    
    pProvData = WTHelperProvDataFromStateData(StateData)
    
    If 0 = pProvData Then Exit Sub
    
    GetMem4 ByVal pProvData, ProvData.cbStruct
    memcpy ProvData, ByVal pProvData, IIf(ProvData.cbStruct < LenB(ProvData), ProvData.cbStruct, LenB(ProvData))    'Win7+ size of struct > &H80
    
    NumberOfSignatures = 0
    Do
        pCPSigner = WTHelperGetProvSignerFromChain(pProvData, NumberOfSignatures, 0, 0)
        
        If 0 <> pCPSigner Then
            
            memcpy CPSigner, ByVal pCPSigner, LenB(CPSigner)
            
            ' count of CRYPT_PROVIDER_CERT structures
            If 0 <> CPSigner.csCertChain And 0 <> CPSigner.pasCertChain Then
                
                'CPSigner.pasCertChain - contains certificates of all chain. Last index is a root cert
                
                ReDim CPCERT(CPSigner.csCertChain - 1)
                
                GetMem4 ByVal ArrPtr(CPCERT()), lpSA
                GetMem4 ByVal lpSA + 12, lpOldPt
                GetMem4 CPSigner.pasCertChain, ByVal lpSA + 12
            
                'For i = 0 To CPSigner.csCertChain - 1
                
                'get a root certificate only
                i = CPSigner.csCertChain - 1
                    
                    ReDim Preserve Signatures(0 To NumberOfSignatures)
                    
                    Signatures(NumberOfSignatures) = CertDuplicateCertificateContext(CPCERT(i).pCert)
                    NumberOfSignatures = NumberOfSignatures + 1
                
                'Next
                
                GetMem4 lpOldPt, ByVal lpSA + 12
                
            End If
            
            'NumberOfSignatures = NumberOfSignatures + 1
        End If
    Loop While pCPSigner

    'WINTRUST_Free ProvData.padwTrustStepErrors
    'WINTRUST_Free ProvData.pPDSip
    'WINTRUST_Free ProvData.psPfns
    'WINTRUST_Free pProvData

End Sub

Sub GetSignerNameFromCertificate(pCertificate As Long) 'ptr -> CERT_CONTEXT
    
    Dim Certificate As CERT_CONTEXT
    Dim pCertInfo As Long
    Dim CertInfo As CERT_INFO
    Dim sName As String
    
    If 0 <> pCertificate Then
    
        memcpy Certificate, ByVal pCertificate, LenB(Certificate)
    
        pCertInfo = Certificate.pCertInfo

        If 0 <> pCertInfo Then
    
            memcpy CertInfo, ByVal pCertInfo, LenB(CertInfo)
    
            sName = GetCertNameString(CertInfo.Subject) ' X.500 string
    
            ' RFC2253 - http://www.ietf.org/rfc/rfc2253.txt
            '
            ' CN  = commonName
            ' L   = localityName
            ' ST  = stateOrProvinceName
            ' O   = organizationName
            ' OU  = organizationalUnitName
            ' C   = countryName
            ' STREET = streetAddress
            ' DC  = domainComponent
            ' UID = userid
    
        End If
    
    End If
    
    
End Sub

Function GetCertNameString(Blob As CRYPTOAPI_BLOB) As String

    Dim bufferSize As Long
    Dim sName As String
    
    bufferSize = CertNameToStr(X509_ASN_ENCODING, VarPtr(Blob), CERT_X500_NAME_STR, 0&, 0&)

    If bufferSize Then

        sName = String$(bufferSize, vbNullChar)
    
        CertNameToStr X509_ASN_ENCODING, VarPtr(Blob), CERT_X500_NAME_STR, StrPtr(sName), bufferSize
    
        sName = Left$(sName, lstrlen(StrPtr(sName)))
    
    End If
    
    Debug.Print sName
    
    GetCertNameString = sName
    
End Function

Public Function GetCertString(pCertContext As Long, ID As Long, Optional SubID As Long) As String

    Dim bufSize As Long
    Dim sName As String

    bufSize = CertGetNameString(pCertContext, ID, SubID, 0, 0, 0)
    
    If bufSize Then
        sName = String$(bufSize, vbNullChar)
        CertGetNameString pCertContext, ID, SubID, 0, StrPtr(sName), bufSize
        sName = Left$(sName, lstrlen(StrPtr(sName)))
    End If
    
     GetCertString = sName
End Function

Function ExtractPropertyFromCertificate(pCertContext As Long, ID As Long) As String
    Dim bufSize As Long
    Dim buf()   As Byte
    Dim i       As Long
    Dim Hash    As String

    CertGetCertificateContextProperty pCertContext, ID, 0&, bufSize
    If bufSize Then
        ReDim buf(bufSize - 1)
        Hash = String$(bufSize * 2, vbNullChar)
        If CertGetCertificateContextProperty(pCertContext, ID, buf(0), bufSize) Then
            For i = 0 To bufSize - 1
                Mid(Hash, i * 2 + 1) = Right$("0" & Hex(buf(i)), 2&)
            Next
        End If
    End If
    
    ExtractPropertyFromCertificate = Hash
End Function
