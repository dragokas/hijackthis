Attribute VB_Name = "CertEnumerator"
Option Explicit

' Root certificate hashes enumerator

' Examples:

' Finding a certificate by subject name from Store
' https://msdn.microsoft.com/en-us/library/windows/desktop/aa382038(v=vs.85).aspx

' CertFindCertificateInStore
' CertEnumCertificatesInStore

' Cert Stores - https://msdn.microsoft.com/en-us/library/system.security.cryptography.x509certificates.storename(v=vs.110).aspx
' Enumeration - https://msdn.microsoft.com/en-us/library/windows/desktop/aa382363(v=vs.85).aspx

Private Declare Function CertOpenSystemStore Lib "Crypt32.dll" Alias "CertOpenSystemStoreW" (ByVal hprov As Long, ByVal szSubsystemProtocol As Long) As Long
Private Declare Function CertCloseStore Lib "Crypt32.dll" (ByVal hCertStore As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertEnumCertificatesInStore Lib "Crypt32.dll" (ByVal hCertStore As Long, ByVal pPrevCertContext As Long) As Long

Private Const CERT_HASH_PROP_ID As Long = 3&
Private Const CERT_SIGNATURE_HASH_PROP_ID As Long = 15&
Private Const CERT_SUBJECT_PUBLIC_KEY_MD5_HASH_PROP_ID As Long = 25&

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
    
    StoreName = "Root"
    
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
                
                If InStr(1, CertName, "microsoft", 1) <> 0 Then
                    Print #ff, CertName & ";" & HashCert & ";" & HashSign & ";" & HashPKey
                End If
                
            End If
            
        Loop While pCertContext <> 0
    
        CertCloseStore hStore, 0&
    
    End If

    Close #ff
    
End Sub
