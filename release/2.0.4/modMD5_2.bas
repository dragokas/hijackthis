Attribute VB_Name = "modMD5_2"
Option Explicit
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_CLASS_HASH As Long = 32768

Private Const HP_HASHVAL As Long = 2
Private Const HP_HASHSIZE As Long = 4

Private Const CRYPT_VERIFYCONTEXT = &HF0000000

Private Const PROV_RSA_FULL As Long = 1
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"

Public Function GetFileMD5(sFilename$, Optional lFileSize&) As String
    If Not FileExists(sFilename) Then Exit Function
    On Error Resume Next
    If FileLen(sFilename) = 0 Then
        'speed tweak :) 0-byte file always has the same MD5
        GetFileMD5 = "D41D8CD98F00B204E9800998ECF8427E"
        Exit Function
    End If
    On Error GoTo 0:

    frmMain.shpMD5Background.Visible = True
    frmMain.shpMD5Progress.Width = 15
    frmMain.shpMD5Progress.Visible = True
    frmMain.lblMD5.Caption = "Calculating MD5 checksum of " & sFilename & "..."
    frmMain.lblMD5.Visible = True
    'DoEvents
    UpdateMD5Progress 0, 0
    
    If lFileSize = 0 Then lFileSize = FileLen(sFilename)
    
    Dim sFileContents$
    On Error Resume Next
    UpdateMD5Progress 1, 8
    'Open sFileName For Binary Access Read As #1
    Open sFilename For Binary Access Read Shared As #1
        If Err Then Exit Function
        sFileContents = Input(lFileSize, #1)
    Close #1
    
    Dim hCrypt&, hHash&, uMD5(255) As Byte, lMD5Len&, i%, sMD5$
    UpdateMD5Progress 2, 8
    If CryptAcquireContext(hCrypt, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        UpdateMD5Progress 3, 8
        If CryptCreateHash(hCrypt, ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_MD5, 0, 0, hHash) <> 0 Then
            UpdateMD5Progress 4, 8
            If CryptHashData(hHash, sFileContents, Len(sFileContents), 0) <> 0 Then
                UpdateMD5Progress 5, 8
                If CryptGetHashParam(hHash, HP_HASHSIZE, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
                    UpdateMD5Progress 6, 8
                    lMD5Len = uMD5(0)
                    If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
                        UpdateMD5Progress 7, 8
                        For i = 0 To lMD5Len - 1
                            sMD5 = sMD5 & Right("0" & Hex(uMD5(i)), 2)
                        Next i
                    End If
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hCrypt, 0
        UpdateMD5Progress 8, 8
    End If
    
    If sMD5 <> vbNullString Then
        GetFileMD5 = " (filesize " & Len(sFileContents) & " bytes, MD5 " & sMD5 & ")"
    End If
    UpdateMD5Progress -1, 0
End Function

Public Function GetFileFromAutostart$(sAutostart$, Optional bGetMD5 As Boolean = True)
    Dim sDummy$
    On Error GoTo Error:
    
    If InStr(sAutostart, "(file missing)") > 0 Then Exit Function
    sDummy = sAutostart
    If FileExists(sDummy) Then
        GetFileFromAutostart = sDummy
        Exit Function
    End If
    'forms we can find the file in:
    'c:\bla\bla.exe
    'c:\bla.exe
    'bla.exe
    'bla
    '
    'also possible:
    '* surrounding quotes
    '* arguments (possibly files)
    
    If Left(sDummy, 1) = """" Then
        'has quotes
        'stripping like this also removes any
        'arguments, so a path means it's finished
        sDummy = Mid(sDummy, 2)
        sDummy = Left(sDummy, InStr(sDummy, """") - 1)
        
        If InStr(sDummy, "\") > 0 Then
            GoTo GetMD5:
        Else
            GoTo FindFullPath:
        End If
    End If
    
    If FileExists(sDummy) Then GoTo GetMD5
    
    If LCase(Right(sDummy, 4)) <> ".exe" And _
       LCase(Right(sDummy, 4)) <> ".com" Then
        'has arguments, or no extension
        If InStr(sDummy, " ") = 0 Then
            'only one word, so no extension
            sDummy = GetLongPath(sDummy & ".exe")
            If InStr(sDummy, "\") = 0 Then
                sDummy = GetLongPath(sDummy & ".com")
            End If
            GoTo GetMD5:
        Else
            'multiple words, the first is the program
            If FileExists(Left(sDummy, InStr(sDummy, " ") - 1)) Then
                sDummy = Left(sDummy, InStr(sDummy, " ") - 1)
                sDummy = GetLongPath(sDummy)
            Else
                sDummy = Left(sDummy, InStrRev(sDummy, " ") - 1)
                sDummy = GetLongPath(sDummy)
            End If
            GoTo GetMD5:
        End If
    End If
    
FindFullPath:
    If InStr(sDummy, "\") = 0 Then
        'no path - so search for file
        sDummy = GetLongPath(sDummy)
    End If
    
    
GetMD5:
    If FileExists(sDummy) Then
        If bGetMD5 Then
            GetFileFromAutostart = GetFileMD5(sDummy)
        Else
            GetFileFromAutostart = sDummy
        End If
    End If
    
    frmMain.lblMD5.Visible = False
    DoEvents
    Exit Function
    
Error:
    ErrorMsg "modMD5_GetFileFromAutostart", Err.Number, Err.Description, sAutostart
End Function

Public Sub UpdateMD5Progress(i&, iMax&)
    With frmMain
        If i = -1 Then
            'hide md5 progress bar
            .shpMD5Background.Visible = False
            .shpMD5Progress.Width = 15
            .shpMD5Progress.Visible = False
            DoEvents
            Exit Sub
        ElseIf i = 0 Then
            'show + reset md5 progress bar
            .shpMD5Background.Visible = True
            .shpMD5Progress.Visible = True
            .shpMD5Progress.Width = 15
            DoEvents
            Exit Sub
        End If
        
        .shpMD5Progress.Width = .shpMD5Background.Width * (CLng(i) / CLng(iMax))
        
        DoEvents
    End With
End Sub

