Attribute VB_Name = "modInternet"
Option Explicit

Private Const MAX_HOSTNAME_LEN = 132&
Private Const MAX_DOMAIN_NAME_LEN = 132&
Private Const MAX_SCOPE_ID_LEN = 260&

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type IP_ADDR_STRING
    Next As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Private Type FIXED_INFO
    HostName As String * MAX_HOSTNAME_LEN
    DomainName As String * MAX_DOMAIN_NAME_LEN
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId  As String * MAX_SCOPE_ID_LEN
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal InternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

Private Declare Function GetNetworkParams Lib "IPHlpApi.dll" (FixedInfo As Any, pOutBufLen As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)


Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_OVERWRITEPROMPT = &H2

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Const INTERNET_SERVICE_HTTP = 3
Private Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000

Private Const ERROR_BUFFER_OVERFLOW = 111&


Private sUserAgent$
'Private Const sURLUpdate$ = "http://www.spywareinfo.com/~merijn/files/HiJackThis-update.txt"
'Private Const sURLDownload$ = "http://www.spywareinfo.com/~merijn/files/HiJackThis.zip"
Private Const sURLUpdate$ = vbNullString
Private Const sURLDownload$ = vbNullString

Public bDebug As Boolean
Public szResponse As String
Public szSubmitUrl As String
Private sTriageObj() As String

Public Function GetDNS(DnsAdresses() As String) As Boolean
    On Error GoTo ErrorHandler:

    AppendErrorLogCustom "GetDNS - Begin"

    Dim DNS()               As String
    Dim FixedInfoBuffer()   As Byte
    Dim FixedInfo           As FIXED_INFO
    Dim Buffer              As IP_ADDR_STRING
    Dim FixedInfoSize       As Long
    Dim pAddrStr            As Long
    
    ReDim DNS(0) As String
    
    If ERROR_BUFFER_OVERFLOW = GetNetworkParams(ByVal 0&, FixedInfoSize) Then
    
        ReDim FixedInfoBuffer(FixedInfoSize - 1)
       
        If ERROR_SUCCESS = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize) Then
            GetDNS = True
            CopyMemory FixedInfo, FixedInfoBuffer(0), Len(FixedInfo)
            DNS(0) = FixedInfo.DnsServerList.IpAddress
            DNS(0) = Left$(DNS(0), lstrlen(StrPtr(DNS(0))))
            pAddrStr = FixedInfo.DnsServerList.Next
            
            Do While pAddrStr <> 0
                CopyMemory Buffer, ByVal pAddrStr, Len(Buffer)
                ReDim Preserve DNS(UBound(DNS) + 1) As String
                DNS(UBound(DNS)) = Left$(Buffer.IpAddress, lstrlen(StrPtr(Buffer.IpAddress)))
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

Public Sub CheckForUpdate()
    Dim hInternet&, hFile&, sBuffer$, lBufferLen&
    Dim sVer$, sUpdate$, sZipFile$, sThisVersion$
    Dim sProxy$, sFileName$
    
    sThisVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    
    Dim szUpdateUrl As String
    'szUpdateUrl = "http://sourceforge.net/projects/hjt/"
    szUpdateUrl = "http://dragokas.com/tools/HiJackThis.zip"
    
    'If IsOnline Then
        ShellExecute 0&, StrPtr("open"), StrPtr(szUpdateUrl), 0&, 0&, vbNormalFocus
    'Else
    '    MsgBoxW "No Internet Connection Available"
    'End If
End Sub

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

Function GetUrl(szUrl As String) As String
    On Error GoTo ErrorHandler:
    Dim szRequest As String
    Dim xmlhttp As Object
    Dim dataLen As Long
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

    dataLen = Len(szRequest)
    xmlhttp.Open "GET", szUrl, False
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    'xmlhttp.setRequestHeader "User-Agent", "HJT.1.99.2" & "|" & sWinVersion & "|" & sMSIEVersion

    xmlhttp.send "" & szRequest
    'msgboxW szData

    GetUrl = xmlhttp.responseText
    'msgboxW szResponse

    Set xmlhttp = Nothing
    Exit Function

ErrorHandler:
    GetUrl = "HJT_NOT_SUPPORTED"
    ErrorMsg Err, "GetUrl"
    If inIDE Then Stop: Resume Next
End Function

Public Sub ParseHTTPResponse(szResponse As String)
    On Error GoTo ErrorHandler:

    Dim curPos As Long
    Dim startIDPos, endIDPos, startDataPos, endDataPos As Long
    Dim szDataId, szData As String

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
                ShellExecute 0&, StrPtr("open"), StrPtr(szData), 0&, 0&, vbNormalFocus
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
    Dim char As String
    
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
    Dim sPath$, sFileName$, sFilesize$, sMD5$, sItem$()
    If Not FileExists(sFile) Then Exit Sub
    If InStr(sFile, "\") = 0 Then Exit Sub
    'sPath = Left$(sFile, InStrRev(sFile, "\") - 1)
    sFileName = Mid$(sFile, InStrRev(sFile, "\") + 1)
    sFilesize = CStr(FileLen(sFile))
    sMD5 = GetFileMD5(sFile, , True)
    
    ReDim sItem(8)
    sItem(0) = sName     'id to item
    sItem(1) = sFileName 'name
    sItem(2) = sCLSID
    sItem(3) = sFile     'complete path+filename
    sItem(4) = sFileName 'filename
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
    sURL = "http://www.spywareguide.com/report/triage.php"
    sUserAgent = "StartupList v" & App.Major & "." & Format$(App.Minor, "00")
    sPost = Mid$(URLEncode(Join(sTriageObj, "&")), 2)
    If sPost = vbNullString Then Exit Function
    sHeaders = "Accept: text/html,text/plain" & vbCrLf & _
               "Accept-Charset: ISO-8859-1,utf-8" & vbCrLf & _
               "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
               "Content-Length: " & Len(sPost)
    
    hInternet = InternetOpen(sUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hInternet = 0 Then Exit Function

    hConnect = InternetConnect(hInternet, "www.spywareguide.com", 80, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    If hConnect > 0 Then
        hRequest = HttpOpenRequest(hConnect, "POST", "/report/triage.php", "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        If hRequest > 0 Then
            HttpSendRequest hRequest, sHeaders, Len(sHeaders), ByVal sPost, Len(sPost)
            sResponse = vbNullString
            Do
                sBuffer = Space$(1024)
                InternetReadFile hRequest, sBuffer, Len(sBuffer), lBufferLen
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

Public Function DownloadFile(sURL$, sTarget$) As Boolean
    On Error GoTo ErrorHandler:

    Dim hInternet&, hFile&, sBuffer$, sFile$, lBytesRead&
    Dim sUserAgent$, ff%
    DownloadFile = False
    If FileExists(sTarget) Then Exit Function
    sUserAgent = "StartupList v" & StartupListVer '& App.Major & "." & Format(App.Minor, "00") & "." & App.Revision
    
    'Status "Downloading Wintrust library..."
    hInternet = InternetOpen(sUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hInternet Then
        hFile = InternetOpenUrl(hInternet, sURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        If hFile Then
            Do
                sBuffer = Space(16384)
                InternetReadFile hFile, sBuffer, Len(sBuffer), lBytesRead
                sFile = sFile & Left(sBuffer, lBytesRead)
            Loop Until lBytesRead = 0
            InternetCloseHandle hFile
            
            ff = FreeFile()
            
            Open sTarget For Output As #ff
                Print #ff, sFile
            Close #ff
            DownloadFile = True
        Else
            'Unable to connect to the Internet.
            MsgBoxW Translate(1005), vbCritical
        End If
        InternetCloseHandle hInternet
    End If
    'Done.
    Status Translate(1006)
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DownloadFile", "URL:", sURL, "Target:", sTarget
    If inIDE Then Stop: Resume Next
End Function

