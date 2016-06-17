Attribute VB_Name = "modMD5_2"
Option Explicit

Private Const MAX_HASH_FILE_SIZE As Currency = 10485760@ '10 MB. (maximum file size to calculate hash)

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextW" (ByRef phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hprov As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal pCryptHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pcbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData_Array Lib "advapi32.dll" Alias "CryptHashData" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptHashData_Str Lib "advapi32.dll" Alias "CryptHashData" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hprov As Long, ByVal dwFlags As Long) As Long

Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_CLASS_HASH As Long = 32768

Private Const HP_HASHVAL As Long = 2
Private Const HP_HASHSIZE As Long = 4

Private Const CRYPT_VERIFYCONTEXT = &HF0000000

Private Const PROV_RSA_FULL As Long = 1
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"

Public Function GetFileMD5(sFileName$, Optional lFileSize&) As String
    On Error GoTo ErrorHandler:
    
    Dim ff          As Long
    Dim hCrypt      As Long
    Dim hHash       As Long
    Dim uMD5(255)   As Byte
    Dim lMD5Len     As Long
    Dim i           As Long
    Dim sMD5        As String
    Dim aBuf()      As Byte
    Dim OldRedir    As Boolean

    ToggleWow64FSRedirection False, sFileName, OldRedir
    
    If Not OpenW(sFileName, FOR_READ, ff) Then GoTo Finalize
    
    If lFileSize = 0 Then lFileSize = LOFW(ff)
    If lFileSize = 0 Then
        'speed tweak :) 0-byte file always has the same MD5
        GetFileMD5 = " (size: 0 bytes, MD5: D41D8CD98F00B204E9800998ECF8427E)"
        Exit Function
    End If
    If lFileSize > MAX_HASH_FILE_SIZE Then
        GetFileMD5 = " (size: " & lFileSize & " bytes)"
        Exit Function
    End If
    
    ReDim aBuf(lFileSize - 1)
    If ff <> 0 And ff <> -1 Then
      GetW ff, 1&, , VarPtr(aBuf(0)), CLng(lFileSize)
      CloseW ff
    End If
    
    DoEvents
    
    'frmMain.shpMD5Background.Visible = True
    'frmMain.shpMD5Progress.Width = 15
    'frmMain.shpMD5Progress.Visible = True
    frmMain.lblMD5.Caption = "Calculating checksum of " & sFileName & "..."
    'frmMain.lblMD5.Visible = True
    'DoEvents
    'UpdateMD5Progress 0, 0
    '...
    'On Error Resume Next
    'UpdateMD5Progress 1, 8
    
    ToggleWow64FSRedirection True
    
    'UpdateMD5Progress 2, 8
    If CryptAcquireContext(hCrypt, 0&, StrPtr(MS_ENHANCED_PROV), PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        'UpdateMD5Progress 3, 8
        If CryptCreateHash(hCrypt, ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_MD5, 0, 0, hHash) <> 0 Then
            'UpdateMD5Progress 4, 8
            If CryptHashData_Array(hHash, aBuf(0), lFileSize, 0) <> 0 Then
                'UpdateMD5Progress 5, 8
                If CryptGetHashParam(hHash, HP_HASHSIZE, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
                    'UpdateMD5Progress 6, 8
                    lMD5Len = uMD5(0)
                    If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
                        'UpdateMD5Progress 7, 8
                        For i = 0 To lMD5Len - 1
                            sMD5 = sMD5 & Right$("0" & Hex(uMD5(i)), 2)
                        Next i
                    End If
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hCrypt, 0&
        'UpdateMD5Progress 8, 8
    Else
        ErrorMsg err, "modMD5_GetFileMD5", "File: ", sFileName$, "Handle: ", ff, "Size: ", lFileSize
    End If
    
    If Len(sMD5) <> 0 Then
        GetFileMD5 = " (size: " & lFileSize & " bytes, MD5: " & sMD5 & ")"
    Else
        GetFileMD5 = " (size: " & lFileSize & " bytes)"
    End If
    
    DoEvents
    'UpdateMD5Progress -1, 0
Finalize:
    If OldRedir = False Then ToggleWow64FSRedirection False
    frmMain.lblMD5.Caption = ""
    Exit Function
ErrorHandler:
    ErrorMsg err, "modMD5_GetFileMD5", "File: ", sFileName$, "Handle: ", ff, "Size: ", lFileSize
    frmMain.lblMD5.Caption = ""
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFileFromAutostart$(sAutostart$, Optional bGetMD5 As Boolean = True)
    Dim sDummy$
    On Error GoTo ErrorHandler:
    
    If InStr(sAutostart, "(file missing)") > 0 Then Exit Function
    
    sDummy = sAutostart

    'forms we can find the file in:
    'c:\bla\bla.exe
    'c:\bla.exe
    'bla.exe
    'bla
    '
    'also possible:
    '* surrounding quotes
    '* arguments (possibly files)
    
    If Not FileExists(sDummy) Then
      If Left$(sDummy, 1) = """" Then
        'has quotes
        'stripping like this also removes any
        'arguments, so a path means it's finished
        sDummy = Mid$(sDummy, 2)
        sDummy = Left$(sDummy, InStr(sDummy, """") - 1)
        
        If InStr(sDummy, "\") = 0 Then
            'GoTo FindFullPath:
            If InStr(sDummy, "\") = 0 Then
                'no path - so search for file
                sDummy = GetLongPath(sDummy)
            End If
        End If
      End If
    End If
    
    If Not FileExists(sDummy) Then
      If LCase$(Right$(sDummy, 4)) <> ".exe" And _
       LCase$(Right$(sDummy, 4)) <> ".com" Then
        'has arguments, or no extension
        If InStr(sDummy, " ") = 0 Then
            'only one word, so no extension
            sDummy = GetLongPath(sDummy & ".exe")
            If InStr(sDummy, "\") = 0 Then
                sDummy = GetLongPath(sDummy & ".com")
            End If
        Else
            'multiple words, the first is the program
            If FileExists(Left$(sDummy, InStr(sDummy, " ") - 1)) Then
                sDummy = Left$(sDummy, InStr(sDummy, " ") - 1)
                sDummy = GetLongPath(sDummy)
            Else
                sDummy = Left$(sDummy, InStrRev(sDummy, " ") - 1)
                sDummy = GetLongPath(sDummy)
            End If
        End If
      End If
    End If
    
    If FileExists(sDummy) Then
        If bGetMD5 Then
            GetFileFromAutostart = GetFileMD5(sDummy)
        Else
            GetFileFromAutostart = sDummy
        End If
    End If
    
    'frmMain.lblMD5.Visible = False
    DoEvents
    Exit Function
ErrorHandler:
    ErrorMsg err, "modMD5_GetFileFromAutostart", sAutostart
    If inIDE Then Stop: Resume Next
End Function

'Public Sub UpdateMD5Progress(i&, iMax&)
'    With frmMain
'        If i = -1 Then
'            'hide md5 progress bar
'            .shpMD5Background.Visible = False
'            .shpMD5Progress.Width = 15
'            .shpMD5Progress.Visible = False
'            DoEvents
'            Exit Sub
'        ElseIf i = 0 Then
'            'show + reset md5 progress bar
'            .shpMD5Background.Visible = True
'            .shpMD5Progress.Visible = True
'            .shpMD5Progress.Width = 15
'            DoEvents
'            Exit Sub
'        End If
'
'        .shpMD5Progress.Width = .shpMD5Background.Width * (CLng(i) / CLng(iMax))
'
'        DoEvents
'    End With
'End Sub

