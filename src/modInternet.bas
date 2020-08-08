Attribute VB_Name = "modInternet"
'[modInternet.bas]

'
' Internet module by Merijn Bellekom & Alex Dragokas
'

Option Explicit

'Private Const MAX_HOSTNAME_LEN = 132&
'Private Const MAX_DOMAIN_NAME_LEN = 132&
'Private Const MAX_SCOPE_ID_LEN = 260&
'
'Private Type OPENFILENAME
'    lStructSize As Long
'    hWndOwner As Long
'    hInstance As Long
'    lpstrFilter As String
'    lpstrCustomFilter As String
'    nMaxCustFilter As Long
'    nFilterIndex As Long
'    lpstrFile As String
'    nMaxFile As Long
'    lpstrFileTitle As String
'    nMaxFileTitle As Long
'    lpstrInitialDir As String
'    lpstrTitle As String
'    flags As Long
'    nFileOffset As Integer
'    nFileExtension As Integer
'    lpstrDefExt As String
'    lCustData As Long
'    lpfnHook As Long
'    lpTemplateName As String
'End Type
'
'Private Type IP_ADDR_STRING
'    Next As Long
'    IpAddress As String * 16
'    IpMask As String * 16
'    Context As Long
'End Type
'
'Private Type FIXED_INFO
'    HostName As String * MAX_HOSTNAME_LEN
'    DomainName As String * MAX_DOMAIN_NAME_LEN
'    CurrentDnsServer As Long
'    DnsServerList As IP_ADDR_STRING
'    NodeType As Long
'    ScopeId  As String * MAX_SCOPE_ID_LEN
'    EnableRouting As Long
'    EnableProxy As Long
'    EnableDns As Long
'End Type
'Public Enum COMPUTER_NAME_FORMAT
'  ComputerNameNetBIOS
'  ComputerNameDnsHostname
'  ComputerNameDnsDomain
'  ComputerNameDnsFullyQualified
'  ComputerNamePhysicalNetBIOS
'  ComputerNamePhysicalDnsHostname
'  ComputerNamePhysicalDnsDomain
'  ComputerNamePhysicalDnsFullyQualified
'  ComputerNameMax
'End Enum
'Public Enum WinHttpRequestOption
'  WinHttpRequestOption_UserAgentString
'  WinHttpRequestOption_URL
'  WinHttpRequestOption_URLCodePage
'  WinHttpRequestOption_EscapePercentInURL
'  WinHttpRequestOption_SslErrorIgnoreFlags
'  WinHttpRequestOption_SelectCertificate
'  WinHttpRequestOption_EnableRedirects
'  WinHttpRequestOption_UrlEscapeDisable
'  WinHttpRequestOption_UrlEscapeDisableQuery
'  WinHttpRequestOption_SecureProtocols
'  WinHttpRequestOption_EnableTracing
'  WinHttpRequestOption_RevertImpersonationOverSsl
'  WinHttpRequestOption_EnableHttpsToHttpRedirects
'  WinHttpRequestOption_EnablePassportAuthentication
'  WinHttpRequestOption_MaxAutomaticRedirects
'  WinHttpRequestOption_MaxResponseHeaderSize
'  WinHttpRequestOption_MaxResponseDrainSize
'  WinHttpRequestOption_EnableHttp1_1
'  WinHttpRequestOption_EnableCertificateRevocationCheck
'End Enum
'
'Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectW" (ByVal InternetSession As Long, ByVal sServerName As Long, ByVal nServerPort As Integer, ByVal sUsername As Long, ByVal sPassword As Long, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
'Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenW" (ByVal sAgent As Long, ByVal lAccessType As Long, ByVal sProxyName As Long, ByVal sProxyBypass As Long, ByVal lFlags As Long) As Long
'Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlW" (ByVal hInternetSession As Long, ByVal sURL As Long, ByVal sHeaders As Long, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
'Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestW" (ByVal hHttpSession As Long, ByVal sVerb As Long, ByVal sObjectName As Long, ByVal sVersion As Long, ByVal sReferer As Long, lplpszAcceptTypes As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestW" (ByVal hHttpRequest As Long, ByVal sHeaders As Long, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Long
'Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoW" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Long
'Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
'
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
'
'Private Declare Function GetNetworkParams Lib "IPHlpApi.dll" (FixedInfo As Any, pOutBufLen As Long) As Long
'Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetComputerNameEx Lib "kernel32.dll" Alias "GetComputerNameExW" (ByVal NameType As COMPUTER_NAME_FORMAT, ByVal lpBuffer As Long, lpnSize As Long) As Long
Private Declare Function SetComputerNameEx Lib "kernel32.dll" Alias "SetComputerNameExW" (ByVal NameType As COMPUTER_NAME_FORMAT, ByVal lpBuffer As Long) As Long
'
'
'Private Const OFN_HIDEREADONLY = &H4
'Private Const OFN_NONETWORKBUTTON = &H20000
'Private Const OFN_PATHMUSTEXIST = &H800
'Private Const OFN_FILEMUSTEXIST = &H1000
'Private Const OFN_OVERWRITEPROMPT = &H2
'
'Private Const INTERNET_OPEN_TYPE_DIRECT = 1
'Private Const INTERNET_FLAG_RELOAD = &H80000000
'
'Private Const INTERNET_SERVICE_HTTP = 3
'Private Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
'
'Private Const ERROR_BUFFER_OVERFLOW = 111&


'Private Const sURLUpdate$ = "http://www.spywareinfo.com/~merijn/files/HiJackThis-update.txt"
'Private Const sURLDownload$ = "http://www.spywareinfo.com/~merijn/files/HiJackThis.zip"

Private Const sURLUpdate$ = vbNullString
Private Const sURLDownload$ = vbNullString

Public bDebug As Boolean
Public szResponse As String
Public szSubmitUrl As String
Private sTriageObj() As String


Public Function GetCompName(NameType As COMPUTER_NAME_FORMAT) As String
    Dim sName As String
    Dim szName As Long
    sName = String$(MAX_PATH, 0)
    szName = Len(sName)
    If GetComputerNameEx(NameType, StrPtr(sName), szName) Then
        GetCompName = Left$(sName, szName)
    End If
End Function

Public Function SetCompName(NameType As COMPUTER_NAME_FORMAT, sName As String) As Boolean
    SetCompName = SetComputerNameEx(NameType, StrPtr(sName))
End Function

Public Function GetDNS(DnsAdresses() As String) As Boolean
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "GetDNS - Begin"

    Dim DNS()               As String
    Dim FixedInfoBuffer()   As Byte
    Dim FixedInfo           As FIXED_INFO
    Dim Buffer              As IP_ADDR_STRING
    Dim FixedInfoSize       As Long
    Dim pAddrStr            As Long
    Dim cbIpAddress         As Long
    
    ReDim DNS(0) As String
    
    cbIpAddress = UBound(FixedInfo.DnsServerList.IpAddress) + 1
    
    If ERROR_BUFFER_OVERFLOW = GetNetworkParams(ByVal 0&, FixedInfoSize) Then
    
        ReDim FixedInfoBuffer(FixedInfoSize - 1)
       
        If ERROR_SUCCESS = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize) Then
            GetDNS = True
            CopyMemory FixedInfo, FixedInfoBuffer(0), Len(FixedInfo)
            DNS(0) = StringFromPtrA(VarPtr(FixedInfo.DnsServerList.IpAddress(0)))
            pAddrStr = FixedInfo.DnsServerList.Next
            
            Do While pAddrStr <> 0
                CopyMemory Buffer, ByVal pAddrStr, Len(Buffer)
                ReDim Preserve DNS(UBound(DNS) + 1) As String
                DNS(UBound(DNS)) = StringFromPtrA(VarPtr(Buffer.IpAddress(0)))
                pAddrStr = Buffer.Next
            Loop
        End If
    End If
    DnsAdresses = DNS
    
    AppendErrorLogCustom "GetDNS - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetDNS"
    If inIDE Then Stop: Resume Next
End Function

Public Sub CheckForUpdate(bSilentIfNoUpdates As Boolean, bSilentReplace As Boolean, bUseTestVersion As Boolean)
    On Error GoTo ErrorHandler:
    
    'bSilentIfNoUpdates - true, to show MsgBoxes. Usually when you press "Update" button in GUI.
    'bSilentReplace - true, to replace exe file in automatic mode.
    'bUseTestVersion - true, to use test (nightly) build.
    
    Dim sThisVersion$, sNewVersion$, sUpdateUrl$, sUpdVersionURL$, bNoConnection As Boolean
    
    RegSaveProxySettings
    
    sThisVersion = AppVerString
    
    Dim lReturnCode As Long, sErrorMsg As String, bRet As Boolean
    
    'sNewVersion = GetUrl("https://github.com/dragokas/hijackthis/raw/devel/src/HiJackThis-update.txt")
    'sNewVersion = GetUrl("https://raw.githubusercontent.com/dragokas/hijackthis/devel/src/HiJackThis-update.txt")
    
    If bUseTestVersion Then
        bRet = GetUrl2_Str("https://dragokas.com/tools/HiJackThis-update-test.txt", sNewVersion, lReturnCode, sErrorMsg)
    Else
        bRet = GetUrl2_Str("https://dragokas.com/tools/HiJackThis-update.txt", sNewVersion, lReturnCode, sErrorMsg)
    End If
    
    If (Not bRet) Or (Not IsVersion(sNewVersion)) Then
    
        bNoConnection = True
        
        If Not bSilentIfNoUpdates Then
            'Unable to connect to the Internet. Do you want to open download page?
            If MsgBoxW(Translate(1005) & vbCrLf & _
                "(Code: " & lReturnCode & ", msg: " & sErrorMsg & ")" & vbCrLf & vbCrLf & _
                Translate(1015), vbYesNo Or vbExclamation, "HiJackThis") = vbNo Then
                    Exit Sub
            End If
        End If
    Else
        If ConvertVersionToNumber(sThisVersion) >= ConvertVersionToNumber(sNewVersion) Then
            If Not bSilentIfNoUpdates Then
                'You have the most fresh version.
                MsgBoxW Translate(1013), vbInformation, "HiJackThis"
            End If
            Exit Sub
        Else
            If Not bSilentIfNoUpdates Then
                If bSilentReplace Then
                    'Update is available. Do you want to close and update the program?
                    If MsgBoxW(Translate(1014) & vbCrLf & Translate(1028), vbYesNo Or vbInformation, "HiJackThis") = vbNo Then
                        Exit Sub
                    End If
                Else
                    'Update is available. Do you want to open download page?
                    If MsgBoxW(Translate(1014) & vbCrLf & Translate(1015), vbYesNo Or vbInformation, "HiJackThis") = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    If bUseTestVersion Then
        sUpdateUrl = "https://dragokas.com/tools/HiJackThis_test.zip"
    Else
        sUpdateUrl = "https://dragokas.com/tools/HiJackThis.zip"
    End If
    
    'sUpdateUrl = "http://sourceforge.net/projects/hjt/"
    'sUpdateUrl = "https://github.com/dragokas/hijackthis/raw/devel/binary/HiJackThis.exe"
    
    If bSilentReplace And Not (Not bSilentIfNoUpdates And bNoConnection) Then
        If Not bNoConnection Then
        
            If DownloadAndUpdateSelf(sUpdateUrl, bSilentIfNoUpdates) Then
            
                If Not inIDE Then
            
                    frmMain.ReleaseMutex
                    g_NeedTerminate = True
                
                    'relaunch new self
                    Proc.ProcessRun AppPath(True), g_sCommandLine, AppPath(False), 1, True
                End If
            End If
        End If
    Else
        OpenURL sUpdateUrl
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "CheckForUpdate"
    If inIDE Then Stop: Resume Next
End Sub

Function IsVersion(sVer As String) As Boolean
    Dim Ver() As String
    Ver = SplitSafe(sVer, ".")
    If UBound(Ver) = 3 Then
        If IsNumeric(Ver(0)) And _
            IsNumeric(Ver(1)) And _
            IsNumeric(Ver(2)) And _
            IsNumeric(Ver(3)) Then
                IsVersion = True
        End If
    End If
End Function

Public Sub SendData(szUrl As String, szData As String)
    On Error GoTo ErrorHandler
    Dim szRequest As String
    Dim xmlhttp As Object
    Dim dataLen As Long
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

    szRequest = "data=" & URLEncode(szData)

    dataLen = Len(szRequest)
    xmlhttp.Open "POST", szUrl, False
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    'xmlhttp.setRequestHeader "User-Agent", "HJT.1.99.2" & "|" & sWinVersion & "|" & sMSIEVersion

    xmlhttp.send "" & szRequest
    'msgboxW szData

    szResponse = xmlhttp.responseText
    'msgboxW szResponse

    Set xmlhttp = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SendData"
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetUrl(szUrl As String) As String
    On Error Resume Next
    Dim TimeOut As Long
    TimeOut = 2000  'milliseconds
    Dim xhr As Object
    Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
    'https://msdn.microsoft.com/en-us/library/windows/desktop/aa384108(v=vs.85).aspx
    xhr.Option(WinHttpRequestOption_EnableRedirects) = True                'true by default
    xhr.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = True     'false by default
    xhr.SetTimeouts TimeOut, TimeOut, TimeOut, TimeOut
    If xhr Is Nothing Then Set xhr = CreateObject("MSXML2.ServerXMLHTTP")
    If xhr Is Nothing Then Set xhr = CreateObject("Microsoft.XMLHTTP")
    If xhr Is Nothing Then Set xhr = CreateObject("WinHttp.WinHttpRequest")
    xhr.Open "GET", szUrl, False
    xhr.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:60.0) Gecko/20100101 Firefox/60.0"
    xhr.setRequestHeader "Content-Type", "text/css" '"application/x-www-form-urlencoded"
    'xhr.setRequestHeader "Content-Length", Len(sEntityBody)
    On Error GoTo ErrorHandler:
    xhr.send 'URLEncode(sEntityBody)
    GetUrl = xhr.responseText
    Set xhr = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetUrl", szUrl
    If inIDE Then Stop: Resume Next
End Function

Public Function GetUrl2_Str( _
    ByVal sURL As String, _
    ByRef s_outRequest As String, _
    Optional ByRef s_outResultCode As Long, _
    Optional ByRef s_outErrorMsg As String) As Boolean

    Dim b() As Byte

    GetUrl2_Str = GetUrl2(sURL, s_outRequest, b, False, s_outResultCode, s_outErrorMsg)

End Function

Public Function GetUrl2_Arr( _
    ByVal sURL As String, _
    ByRef byte_outRequest() As Byte, _
    Optional ByRef s_outResultCode As Long, _
    Optional ByRef s_outErrorMsg As String) As Boolean
    
    Dim str As String
    
    GetUrl2_Arr = GetUrl2(sURL, str, byte_outRequest, True, s_outResultCode, s_outErrorMsg)
    
End Function

Public Function GetUrl2( _
    ByVal sURL As String, _
    ByRef s_outRequest As String, _
    ByRef byte_outRequest() As Byte, _
    ByVal bUseByteMode As Boolean, _
    ByRef s_outResultCode As Long, _
    ByRef s_outErrorMsg As String) As Boolean
    
    On Error GoTo ErrorHandler:
    
    sURL = NormalizeInetProtocol(sURL)
    
    Dim frm As Form
    Set frm = frmMain
    
    Dim cInet As clsHttpHelps
    Set cInet = New clsHttpHelps
    
    With cInet
        .AutomatiRedirection = True
        .RequestMethod = cGET
        .TimeOut = 5000
        .UseProxy = frm.optProxyManual.Value
        .UseProxyIE = frm.optProxyIE.Value
        
        If .UseProxy Then
            .ProxyAddress = frm.txtUpdateProxyHost.Text & ":" & frm.txtUpdateProxyPort.Text
            .UseProxySocks4 = frm.chkSocks4.Value
        End If
        If .UseProxy Or .UseProxyIE Then
            .UseProxyAuthorization = frm.chkUpdateUseProxyAuth.Value
            
            If .UseProxyAuthorization Then
                .ProxyUser = frm.txtUpdateProxyLogin.Text
                .ProxyPass = frm.txtUpdateProxyPass.Text
            End If
        End If
        .UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:60.0) Gecko/20100101 Firefox/60.0"
        .URL = sURL
        GetUrl2 = .API_Http()
        
        If GetUrl2 Then
            If bUseByteMode Then
                byte_outRequest = .ReturnBytArry
            Else
                s_outRequest = .ReturnString
            End If
        Else
            s_outErrorMsg = .ReturnStatus
        End If
        
        s_outResultCode = .ReturnCode
    End With
    
    Set cInet = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetUrl", sURL
    If inIDE Then Stop: Resume Next
End Function

Public Sub ParseHTTPResponse(szResponse As String)
    On Error GoTo ErrorHandler:

    Dim curPos As Long
    Dim startIDPos As Long, endIDPos As Long, startDataPos As Long, endDataPos As Long
    Dim szDataId As String, szData As String

    curPos = 1
    Do While curPos < Len(szResponse)
        startIDPos = InStr(curPos, szResponse, "#HJT_DATA:", vbTextCompare)
    
        If 1 > startIDPos Then Exit Sub
    
        startIDPos = startIDPos + 10
    
        endIDPos = InStr(curPos, szResponse, "=", vbTextCompare)
    
        If 1 > endIDPos Then Exit Sub
    
        endIDPos = endIDPos
    
        startDataPos = endIDPos + 1
    
        endDataPos = InStr(curPos, szResponse, "#END_HJT_DATA", vbTextCompare)
    
        If 1 > endIDPos Then Exit Sub
    
        endDataPos = endDataPos
    
        curPos = curPos + endDataPos + 14
    
        szDataId = Mid$(szResponse, startIDPos, endIDPos - startIDPos)
        szData = Mid$(szResponse, startDataPos, endDataPos - startDataPos)
    
        Select Case szDataId
            Case "REPORT_URL"
                'ShellExecute 0&, StrPtr("open"), StrPtr(szData), 0&, 0&, vbNormalFocus
            Case "SUBMIT_URL"
                szSubmitUrl = szData
        End Select
    Loop
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "ParseHTTPResponse", "Response:", szResponse
    If inIDE Then Stop: Resume Next
End Sub

Function URLEncode(ByVal Text As String) As String
    On Error GoTo ErrorHandler:

    Dim i As Long
    Dim acode As Long
    
    URLEncode = Text
    
    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid$(URLEncode, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars
            Case 32
                ' replace space with "+"
                Mid$(URLEncode, i, 1) = "+"
            Case Else
                ' replace punctuation chars with "%hex"
                URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$ _
                    (URLEncode, i + 1)
        End Select
    Next
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "URLEncode", "Src:", Text
    If inIDE Then Stop: Resume Next
End Function

Public Function IsOnline() As Boolean

   IsOnline = InternetGetConnectedState(0&, 0&)
     
End Function

' ---------------------------------------------------------------------------------------------------
' StartupList2 routine
' ---------------------------------------------------------------------------------------------------

Public Sub AddTriageObj(sName$, sType$, sFile$, Optional sCLSID$, Optional sCodebase$)
    Dim sFilename$, sFilesize$, sMD5$, sItem$()
    If Not FileExists(sFile) Then Exit Sub
    If InStr(sFile, "\") = 0 Then Exit Sub
    'sPath = Left$(sFile, InStrRev(sFile, "\") - 1)
    sFilename = Mid$(sFile, InStrRev(sFile, "\") + 1)
    sFilesize = CStr(FileLen(sFile))
    sMD5 = GetFileCheckSum(sFile, , True)
    
    ReDim sItem(8)
    sItem(0) = sName     'id to item
    sItem(1) = sFilename 'name
    sItem(2) = sCLSID
    sItem(3) = sFile     'complete path+filename
    sItem(4) = sFilename 'filename
    sItem(5) = sFilesize
    sItem(6) = sMD5
    sItem(7) = sType
    sItem(8) = sCodebase 'Codebase, for DPF
    
    On Error Resume Next
    If UBound(sTriageObj) = -2 Then ReDim sTriageObj(0)
    If Err Then ReDim sTriageObj(0)
    On Error GoTo 0:
    
    ReDim Preserve sTriageObj(UBound(sTriageObj) + 1)
    sTriageObj(UBound(sTriageObj)) = "ITEM[]=" & Join(sItem, "|")
End Sub

Public Function GetTriage$()
    Dim hInternet&, hConnect&, sURL$, sUserAgent$, sPost$
    Dim hRequest&, sResponse$, sBuffer$, lBufferLen&, sHeaders$
    sURL = "https://www.spywareguide.com/report/triage.php"
    sUserAgent = "StartupList v" & App.Major & "." & Format$(App.Minor, "00")
    sPost = Mid$(URLEncode(Join(sTriageObj, "&")), 2)
    If sPost = vbNullString Then Exit Function
    sHeaders = "Accept: text/html,text/plain" & vbCrLf & _
               "Accept-Charset: ISO-8859-1,utf-8" & vbCrLf & _
               "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
               "Content-Length: " & Len(sPost)
    
    hInternet = InternetOpen(StrPtr(sUserAgent), INTERNET_OPEN_TYPE_DIRECT, 0&, 0&, 0)
    If hInternet = 0 Then Exit Function

    hConnect = InternetConnect(hInternet, StrPtr("www.spywareguide.com"), 80, 0&, 0&, INTERNET_SERVICE_HTTP, 0, 0)
    If hConnect > 0 Then
        hRequest = HttpOpenRequest(hConnect, StrPtr("POST"), StrPtr("/report/triage.php"), StrPtr("HTTP/1.1"), 0&, ByVal 0, INTERNET_FLAG_RELOAD, 0)
        If hRequest > 0 Then
            HttpSendRequest hRequest, StrPtr(sHeaders), Len(sHeaders), ByVal StrPtr(sPost), Len(sPost)
            sResponse = vbNullString
            Do
                sBuffer = Space$(1024)
                InternetReadFileString hRequest, sBuffer, Len(sBuffer), lBufferLen
                sBuffer = Left$(sBuffer, lBufferLen)
                sResponse = sResponse & sBuffer
            Loop Until lBufferLen = 0
            GetTriage = sResponse
            InternetCloseHandle hRequest
        End If
        InternetCloseHandle hConnect
    End If
    InternetCloseHandle hInternet
End Function

Public Function DownloadFile(sURL$, sTarget$, Optional bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler:

    Const Chunk As Long = 16384

    Dim hInternet&, hFile&, sFile$, lBytesRead&
    Dim sUserAgent$, ff%
    Dim aBuf() As Byte, curPos As Long
    
    sUserAgent = "StartupList v" & StartupListVer
    
    hInternet = InternetOpen(StrPtr(sUserAgent), INTERNET_OPEN_TYPE_DIRECT, 0&, 0&, 0&)
    
    If hInternet Then
        hFile = InternetOpenUrl(hInternet, StrPtr(sURL), 0&, 0&, INTERNET_FLAG_RELOAD, 0&)
        
        If hFile = 0 And InStr(1, sURL, "https://", 1) <> 0 And OSver.MajorMinor <= 5.2 Then 'XP + https ?
            hFile = InternetOpenUrl(hInternet, StrPtr(Replace$(sURL, "https://", "http://", , , 1)), 0&, 0&, INTERNET_FLAG_RELOAD, 0&)
        End If
        
        If hFile <> 0 Then
            curPos = -1
            DownloadFile = True
            Do
                ReDim Preserve aBuf(curPos + Chunk)
                InternetReadFile hFile, aBuf(curPos + 1), Chunk, lBytesRead
                If lBytesRead < Chunk Then
                    If curPos + lBytesRead <> -1 Then
                        ReDim Preserve aBuf(curPos + lBytesRead)
                        DownloadFile = True
                    End If
                    Exit Do
                Else
                    curPos = curPos + Chunk
                End If
            Loop Until lBytesRead = 0
            
            InternetCloseHandle hFile
            
            If DownloadFile Then
                ff = FreeFile()
                If FileExists(sTarget) Then DeleteFileW StrPtr(sTarget)
                Open sTarget For Binary Access Write As #ff
                    Put #ff, , aBuf
                Close #ff
            End If
        Else
            If Not bSilent Then
                'Unable to connect to the Internet.
                MsgBoxW Translate(1005), vbCritical
            End If
        End If
        InternetCloseHandle hInternet
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DownloadFile", "URL:", sURL, "Target:", sTarget
    DownloadFile = False
    If inIDE Then Stop: Resume Next
End Function

Public Function DownloadFileToArray(sURL$, aBuf() As Byte, Optional bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler:

    Const Chunk As Long = 16384

    Dim hInternet&, hFile&, sFile$, lBytesRead&
    Dim sUserAgent$, ff%
    Dim curPos As Long
    
    sUserAgent = "StartupList v" & StartupListVer
    
    hInternet = InternetOpen(StrPtr(sUserAgent), INTERNET_OPEN_TYPE_DIRECT, 0&, 0&, 0&)
    
    If hInternet Then
        hFile = InternetOpenUrl(hInternet, StrPtr(sURL), 0&, 0&, INTERNET_FLAG_RELOAD, 0&)
        
        If hFile = 0 And InStr(1, sURL, "https://", 1) <> 0 And OSver.MajorMinor <= 5.2 Then 'XP + https ?
            hFile = InternetOpenUrl(hInternet, StrPtr(Replace$(sURL, "https://", "http://", , , 1)), 0&, 0&, INTERNET_FLAG_RELOAD, 0&)
        End If
        
        If hFile <> 0 Then
            curPos = -1
            Do
                ReDim Preserve aBuf(curPos + Chunk)
                InternetReadFile hFile, aBuf(curPos + 1), Chunk, lBytesRead
                If lBytesRead < Chunk Then
                    If curPos + lBytesRead <> -1 Then
                        ReDim Preserve aBuf(curPos + lBytesRead)
                        DownloadFileToArray = True
                    Else
                        Erase aBuf
                    End If
                    Exit Do
                Else
                    curPos = curPos + Chunk
                End If
            Loop Until lBytesRead = 0
            
            InternetCloseHandle hFile
        Else
            If Not bSilent Then
                'Unable to connect to the Internet.
                MsgBoxW Translate(1005), vbCritical
            End If
        End If
        InternetCloseHandle hInternet
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DownloadFile", "URL:", sURL
    DownloadFileToArray = False
    If inIDE Then Stop: Resume Next
End Function

Public Function OpenURL(sEnglishURL As String, Optional sRussianURL As String, Optional bCheckLangByCurrentSelected As Boolean = False) As Boolean
    'by default, language has checked by OS interface
    Dim szUrl As String
    Dim szDefault As String
    
    szDefault = sEnglishURL
    
    If bCheckLangByCurrentSelected Then
        If g_CurrentLang = "Russian" Then
            szUrl = sRussianURL
        Else
            szUrl = sEnglishURL
        End If
    
    ElseIf (IsRussianLangCode(OSver.LangSystemCode) Or IsRussianLangCode(OSver.LangDisplayCode)) And Not (bForceEN Or bForceFR) Then
        szUrl = sRussianURL
    Else
        szUrl = sEnglishURL
    End If
    
    If szUrl = "" Then szUrl = szDefault
    
    If StrBeginWith(szUrl, "https") And OSver.MajorMinor <= 5.2 Then
        szUrl = Replace$(szUrl, "https", "http", 1, 1, 1)
    End If
    
    '// TODO: run ShellExecute with non-Elevated privilages
    OpenURL = (32 < ShellExecute(0&, StrPtr("open"), StrPtr(szUrl), 0&, 0&, vbNormalFocus))
End Function

Public Function DownloadUnzipAndRun(ZipURL As String, FileName As String, bSilent As Boolean) As Boolean
    Dim ArcPath As String
    Dim ExePath As String
    Dim bRun As Boolean
    
    ArcPath = GetEmptyName(BuildPath(TempCU, GetFileName(Replace$(ZipURL, "/", "\"), True)))
    ExePath = BuildPath(AppPath(False), FileName)
    
    If Not bSilent Then
        'Download the program via Internet?
        If MsgBoxW(Translate(1024), vbYesNo Or vbQuestion, GetFileName(FileName)) = vbNo Then Exit Function
    End If
    
    If DownloadFile(ZipURL, ArcPath, True) Or FileExists(ArcPath) Then
        UnpackZIP ArcPath, AppPath(False)
        'Downloading is completed. Run the program?
        If Not bSilent Then
            bRun = MsgBoxW(Translate(1026), vbYesNo, GetFileName(FileName)) = vbYes
        End If
        If bRun Or bSilent Then
            DownloadUnzipAndRun = Proc.ProcessRun(ExePath, "", AppPath(False), 1, True)
        End If
    Else
        'Downloading is failed -> trying to open link in default browser
        '// TODO: run ShellExecute with non-Elevated privilages
        ShellExecute g_HwndMain, StrPtr("open"), StrPtr(ZipURL), 0&, 0&, 1
    End If
End Function

Private Function DownloadAndUpdateSelf(ZipURL As String, bSilent As Boolean) As Boolean
    
    Dim ArcPath     As String
    Dim ExePath     As String
    Dim SignResult  As SignResult_TYPE
    Dim bData()     As Byte
    Dim bDownloaded As Boolean
    Dim hFile       As Long
    
    ArcPath = GetEmptyName(BuildPath(TempCU, GetFileName(Replace$(ZipURL, "/", "\"), True)))
    ExePath = BuildPath(TempCU, "HiJackThis.exe")
    
    'If DownloadFile(ZipURL, ArcPath, True) Then bDownloaded = True
    
    'proxy support
    If GetUrl2_Arr(ZipURL, bData) Then
        If AryPtr(bData) Then
        
            If OpenW(ArcPath, FOR_OVERWRITE_CREATE, hFile) Then
                PutW hFile, 1, VarPtr(bData(0)), UBound(bData) + 1, False
                CloseW hFile
                bDownloaded = True
            Else
                If Not bSilent Then
                    MsgBoxW "Cannot open file to write: " & ArcPath, vbExclamation, "HiJackThis"
                End If
            End If
        End If
    End If
    
    If bDownloaded Then
        UnpackZIP ArcPath, GetParentDir(ArcPath)
        
        If FileExists(ExePath) Then
        
            'checking digital signature
            SignVerify ExePath, SV_PreferInternalSign, SignResult
            
            If IsDragokasSign(SignResult) Then
                
                If FileExists(AppPath(True) & ".bak") Then DeleteFileWEx StrPtr(AppPath(True) & ".bak"), , True
                
                'replacing ...
                'move self
                If 0 = MoveFile(StrPtr(AppPath(True)), StrPtr(AppPath(True) & ".bak")) Then
                
                    'if failed, use cmd.exe method
                    
                    frmMain.ReleaseMutex
                    
                    Proc.ProcessRun _
                        Environ("ComSpec"), _
                        "/d /c (cd\& for /L %+ in (1,1,10) do ((timeout /t 1|| ping 127.1 -n 2)& " & _
                        "move /y """ & AppPath(True) & """" & " " & """" & AppPath(True) & ".bak" & """ && " & _
                        "move /y """ & ExePath & """ """ & AppPath(True) & """ && " & _
                        "start """" """ & AppPath(True) & """ " & g_sCommandLine & "&& exit))", _
                        SysDisk, vbHide, True
                    
                    g_NeedTerminate = True
                Else
                    'move new
                    If 0 <> MoveFile(StrPtr(ExePath), StrPtr(AppPath(True))) Then
                        DownloadAndUpdateSelf = True
                    Else
                        If Not bSilent Then
                            MsgBoxW "Cannot move updated file on self!", vbExclamation, "HiJackThis"
                        End If
                        'revert own filename
                        Call MoveFile(StrPtr(AppPath(True) & ".bak"), StrPtr(AppPath(True)))
                    End If
                End If
            Else
                If Not bSilent Then
                    MsgBoxW "Unpacked file is damaged! Try update again.", vbExclamation, "HiJackThis"
                End If
            End If
        Else
            If Not bSilent Then
                MsgBoxW "Cannot unpack the update! Try again.", vbExclamation, "HiJackThis"
            End If
        End If
    Else
        If Not bSilent Then
            MsgBoxW "Cannot download the update! Try again.", vbExclamation, "HiJackThis"
        End If
    End If
    
    'clear
    If FileExists(ArcPath) Then DeleteFileWEx StrPtr(ArcPath)
    
End Function

Public Function IsDragokasSign(SignResult As SignResult_TYPE) As Boolean

    If (SignResult.isSelfSigned And StrComp(SignResult.HashRootCert, "05F1F2D5BA84CDD6866B37AB342969515E3D912E", 1) = 0) Then
        IsDragokasSign = True
    ElseIf (SignResult.isLegit) Then
        If Date < #7/24/2023# Then
            If StrComp(SignResult.HashFinalCert, "1B78EF517E81A07D1C1C4C6ADFA66A2B7C3269C3", 1) = 0 Then
                IsDragokasSign = True
            End If
        Else
            If InStr(1, SignResult.SubjectName, "Stanislav Polshyn", 1) <> 0 Then
                IsDragokasSign = True
            End If
        End If
    End If
End Function

Public Function NormalizeInetProtocol(ByVal sURL As String) As String

    If OSver.MajorMinor >= 6 Then
        If StrBeginWith(sURL, "http://") Then sURL = Replace$(sURL, "http://", "https://", 1, 1, 1)
    Else
        If StrBeginWith(sURL, "https://") Then sURL = Replace$(sURL, "https://", "http://", 1, 1, 1)
    End If
    NormalizeInetProtocol = sURL
End Function

Public Sub RegSaveProxySettings()
    RegSaveHJT "ProxyServer", frmMain.txtUpdateProxyHost.Text
    RegSaveHJT "ProxyPort", frmMain.txtUpdateProxyPort.Text
    RegSaveHJT "ProxyLogin", frmMain.txtUpdateProxyLogin.Text
    RegSaveHJT "ProxyPass", Crypt(frmMain.txtUpdateProxyPass.Text)
End Sub
