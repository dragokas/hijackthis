Attribute VB_Name = "CertEnumerator"
Option Explicit

' Root certificates hashes enumerator by Alex Dragokas

' Examples:

' Finding a certificate by subject name from Store
' https://msdn.microsoft.com/en-us/library/windows/desktop/aa382038(v=vs.85).aspx

' Cert Stores - https://msdn.microsoft.com/en-us/library/system.security.cryptography.x509certificates.storename(v=vs.110).aspx
' Enumeration - https://msdn.microsoft.com/en-us/library/windows/desktop/aa382363(v=vs.85).aspx

Const MicrosoftOnly = False

Private Declare Function CertOpenSystemStore Lib "Crypt32.dll" Alias "CertOpenSystemStoreW" (ByVal hprov As Long, ByVal szSubsystemProtocol As Long) As Long
Private Declare Function CertCloseStore Lib "Crypt32.dll" (ByVal hCertStore As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertEnumCertificatesInStore Lib "Crypt32.dll" (ByVal hCertStore As Long, ByVal pPrevCertContext As Long) As Long
Private Declare Function CertGetCertificateContextProperty Lib "Crypt32.dll" (ByVal pCertContext As Long, ByVal dwPropId As Long, pvData As Any, pcbData As Long) As Long
Private Declare Function CertGetNameString Lib "Crypt32.dll" Alias "CertGetNameStringW" (ByVal pCertContext As Long, ByVal dwType As Long, ByVal dwFlags As Long, pvTypePara As Any, ByVal pszNameString As Long, ByVal cchNameString As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Const CERT_HASH_PROP_ID As Long = 3&                            'hash of whole certificate
Private Const CERT_SIGNATURE_HASH_PROP_ID As Long = 15&                 'hash of signature of certificate
Private Const CERT_SUBJECT_PUBLIC_KEY_MD5_HASH_PROP_ID As Long = 25&    'MD5 hash of public key

Private Const CERT_NAME_SIMPLE_DISPLAY_TYPE As Long = 4&
Private Const CERT_NAME_ISSUER_FLAG As Long = 1&


Public Sub EnumCertificates()
    Dim StoreName As String
    Dim hStore As Long
    Dim pCertContext As Long
    Dim ff As Integer
    Dim CertName As String
    Dim HashCert As String
    Dim HashSign As String
    Dim HashPKey As String
    
    ff = FreeFile(): Open App.Path & "\" & "Hashes.csv" For Output As #ff
    Print #ff, "Certificate name;Hash cert;Hash Sign;Hash pub.key"
    
    StoreName = "Disallowed"
    
    hStore = CertOpenSystemStore(0, StrPtr(StoreName))

    If 0 <> hStore Then
        Do
            pCertContext = CertEnumCertificatesInStore(hStore, pCertContext)
            
            If 0 <> pCertContext Then
                
                CertName = GetCertString(pCertContext, CERT_NAME_SIMPLE_DISPLAY_TYPE)
                
                Debug.Print "Certificate for: " & CertName
                
                HashCert = ExtractPropertyFromCertificate(pCertContext, CERT_HASH_PROP_ID)
                HashSign = ExtractPropertyFromCertificate(pCertContext, CERT_SIGNATURE_HASH_PROP_ID)
                HashPKey = ExtractPropertyFromCertificate(pCertContext, CERT_SUBJECT_PUBLIC_KEY_MD5_HASH_PROP_ID)
                
                Debug.Print "Hash -> Cert         -> " & HashCert
                Debug.Print "Hash -> Cert Sign    -> " & HashSign
                Debug.Print "Hash -> Cert pub.key -> " & HashPKey
                
                If (MicrosoftOnly And InStr(1, CertName, "microsoft", 1) <> 0) Or Not MicrosoftOnly Then
                    Print #ff, CertName & ";" & HashCert & ";" & HashSign & ";" & HashPKey
                End If
                
            End If
            
        Loop While pCertContext <> 0
    
        CertCloseStore hStore, 0&
    
    End If

    Close #ff
    
End Sub

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
