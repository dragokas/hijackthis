VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHttpHelps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsHttpHelps.cls]

'
' "Http Helps" class by xxdoc123
'
' Fork by Alex Dragokas
'
' v.1.3
'  - Added Socks4 proxy support
'  - Added some new TimeOut options (private)
'
' v.1.2
'  - Proxy is now apllied to all protocols (by default it was http only)
'  - Fixed "Out of range" error when no data is read to m_BytReturnBytArry()
'  - Improved retrieving of status message
'  - Added error 400 as a mark for .API_Http to return 'false'
'  - Added property .UseProxyIE to use system (Internet Explorer) setting of host/port proxy.
'
' v.1.1
'  - Converted functions into Unicode, fixed types. WARNING: it doesn't mean these functions can fully support Unicode (that is WinInet restriction) !!!
'  - Added ReturnCode / ReturnString properties.
'  - API_Http returns 'false' on 404 error now
'  - Fixed 3 potential infinite loops (cert / proxy / auth. requests resending)
'  - Changed default 'UserAgent' into 'Firefox'
'  - Added error handlers, modified return status text logic
'  - 'm_bAutomatiRedirection' bool meaning is set vice versa: true - to follow redirects
'  - Made all arguments of 'API_Http' function as optional. They can be set by class properties
'  - Removed unused (garbage) properties
'  - Renamed some properties

Option Explicit

Enum Method
    [cGET] = 1
    [cPOST] = 2
End Enum

Const CHUNK_SIZE As Long = &H2000&

Private Const INTERNET_COOKIE_HTTPONLY As Long = 8192& 'INTERNET_COOKIE_HTTPONLY = 8192, //Requires IE 8 or higher

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenW" (ByVal sAgent As Long, ByVal lAccessType As Long, ByVal sProxyName As Long, ByVal sProxyBypass As Long, ByVal lFlags As Long) As Long
' Queries an Internet option on the specified handle
Private Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionW" (ByVal hInternet As Long, ByVal dwOption As Long, ByVal lpBuffer As Long, lpdwBufferLength As Long) As Long
' Reads data from a handle opened by the HttpOpenRequest function.
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, Buffer As Any, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
' Closes a single Internet handle or a subtree of Internet handles.
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
' Opens an HTTP request handle.
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestW" (ByVal hHttpSession As Long, ByVal sVerb As Long, ByVal sObjectName As Long, ByVal sVersion As Long, ByVal sReferer As Long, lplpszAcceptTypes As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
' Sends the specified request to the HTTP server.
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestW" (ByVal hRequest As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal lpOptional As Long, ByVal dwOptionalLength As Long) As Long
' Adds one or more HTTP request headers to the HTTP request handle.
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersW" (ByVal hRequest As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal dwModifiers As Long) As Long
Private Declare Function InternetSetCookie Lib "wininet.dll" Alias "InternetSetCookieW" (ByVal lpszUrl As Long, ByVal lpszCookieName As Long, ByVal lpszCookieData As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectW" (ByVal InternetSession As Long, ByVal sServerName As Long, ByVal nServerPort As Integer, ByVal sUsername As Long, ByVal sPassword As Long, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionW" (ByVal hInternet As Long, ByVal dwOption As Long, ByVal lpBuffer As Long, ByVal dwBufferLength As Long) As Long
' Queries for information about an HTTP request.
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoW" (ByVal hRequest As Long, ByVal dwInfoLevel As Long, ByVal lpvBuffer As Long, ByVal lpdwBufferLength As Long, ByVal lpdwIndex As Long) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetGetCookieEx Lib "wininet.dll" Alias "InternetGetCookieExA" (ByVal URL As String, ByVal cookieName As String, ByVal cookieData As String, ByRef Size As Long, ByVal Flags As Long, ByVal pReserved As Long) As Long
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwnd As Long, ByVal sConnectoid As String, ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetCheckConnectionW Lib "wininet.dll" (ByVal lpszUrl As Long, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetOpenUrlW Lib "wininet.dll" (ByVal hInternet As Long, ByVal lpszUrl As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Arguments As Any) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long

Private Const INTERNET_CONNECTION_MODEM      As Long = &H1
Private Const INTERNET_CONNECTION_LAN        As Long = &H2
Private Const INTERNET_CONNECTION_PROXY      As Long = &H4
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_RAS_INSTALLED         As Long = &H10
Private Const INTERNET_CONNECTION_OFFLINE    As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
Private Const INTERNET_COOKIE_THIRD_PARTY    As Long = 131072
Private Const INTERNET_FLAG_RESTRICTED_ZONE  As Long = 16

'Private inIDE As Boolean

' Variable to hold 'RequestHeaders' property value
Private m_StrRequestHeaders As String
' Variable to hold 'RequestCookies' property value
Private m_StrRequestCookies As String
' Variable to hold 'RequestMethod' property value
Private m_StrRequestMethod  As Method
' Variable to hold 'URL' property value
Private m_StrURL            As String
' post data
Private m_VarPostDatas      As Variant
' Variable to hold 'Referer' property value
Private m_StrReferer        As String
' Variable to hold 'ProxyIp' property value
Private m_StrProxyIp        As String
' Variable to hold 'Timeout' property value
Private m_LonTimeout        As Long

Const INTERNET_DEFAULT_FTP_PORT = 21
Const INTERNET_DEFAULT_GOPHER_PORT = 70
Const INTERNET_DEFAULT_HTTP_PORT = 80
Const INTERNET_DEFAULT_HTTPS_PORT = 443
Const INTERNET_DEFAULT_SOCKS_PORT = 1080
Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Const INTERNET_OPTION_SEND_TIMEOUT = 5
Const INTERNET_OPTION_DATA_RECEIVE_TIMEOUT = 8
Const INTERNET_OPTION_DATA_SEND_TIMEOUT = 7
Const INTERNET_OPTION_DISCONNECTED_TIMEOUT = 49 'Not impl
Const INTERNET_OPTION_FROM_CACHE_TIMEOUT = 63
Const INTERNET_OPTION_LISTEN_TIMEOUT = 11 'Not Impl
Const INTERNET_REQFLAG_NET_TIMEOUT = &H80&

Const INTERNET_OPTION_USERNAME = 28
Const INTERNET_OPTION_PASSWORD = 29
Const INTERNET_OPTION_PROXY_USERNAME = 43
Const INTERNET_OPTION_PROXY_PASSWORD = 44
' Type of service to access.
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_SERVICE_GOPHER = 2
Const INTERNET_SERVICE_HTTP = 3

' Brings the data across the wire even if it locally cached.
Const INTERNET_FLAG_RELOAD = &H80000000
Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
' Security constants
Const INTERNET_OPTION_SECURITY_FLAGS = 31
Const SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100
Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
Const INTERNET_FLAG_SECURE = &H800000

' (HttpQueryInfo) The possible values for the lInfoLevel parameter include:
Const HTTP_QUERY_CONTENT_TYPE = 1
Const HTTP_QUERY_CONTENT_LENGTH = 5
Const HTTP_QUERY_EXPIRES = 10
Const HTTP_QUERY_LAST_MODIFIED = 11
Const HTTP_QUERY_PRAGMA = 17
Const HTTP_QUERY_VERSION = 18
Const INTERNET_FLAG_NO_COOKIES = 524288
Const HTTP_QUERY_STATUS_CODE = 19
Const HTTP_QUERY_STATUS_TEXT = 20
Const HTTP_QUERY_RAW_HEADERS = 21
Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Const HTTP_QUERY_FORWARDED = 30
Const HTTP_QUERY_SERVER = 37
Const HTTP_QUERY_USER_AGENT = 39
Const HTTP_QUERY_SET_COOKIE = 43
Const HTTP_QUERY_REQUEST_METHOD = 45
Const HTTP_STATUS_DENIED = 401
Const HTTP_STATUS_PROXY_AUTH_REQ = 407
' Add this flag to the about flags to get request header.
Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Const HTTP_QUERY_FLAG_NUMBER = &H20000000

' Returns the version number of Wininet.dll.
Const INTERNET_OPTION_VERSION = 40
' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Private Type tWinInetDLLVersion
    lMajorVersion                                 As Long
    lMinorVersion                                 As Long
End Type

' (HttpAddRequestHeaders)
' Flags to modify the semantics of this function. Can be a combination of these values:
' Adds the header only if it does not already exist; otherwise, an error is returned.
Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000
' Adds the header if it does not exist. Used with REPLACE.
Const HTTP_ADDREQ_FLAG_ADD = &H20000000
' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

Const FLAG_ICC_FORCE_CONNECTION = &H1
Const INTERNET_OPEN_TYPE_DIRECT = &H1
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_OPEN_TYPE_PROXY = 3

Const INTERNET_FLAG_NO_AUTO_REDIRECT = 2097152
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000

Private Const ERROR_INSUFFICIENT_BUFFER       As Long = 122
'Public ReturnHeads                                 As String
Public m_ReturnStatus                                  As String
'Public ReturnCookies                               As String
Public Event RequestFinished(URL As String, ReturBytArry() As Byte, ByVal lResponseLen As Long)
Public Event RequestError(URL As String, ByVal ErrScriping As String)
Dim hInternetSession As Long
Public Event RequestHeads(URL As String, Heads As String)
' Variable to hold 'UserAgent' property value
Private m_StrsUserAgent    As String
' Variable to hold 'UseProxy' property value
Private m_bbUseProxy       As Boolean
' Variable to hold 'UseProxyIE' property value
Private m_bbUseProxyIE     As Boolean
' Variable to hold 'sProxyAddress' property value
Private m_bbUseProxySocks4 As Boolean
' Variable to hold 'UseProxySocks4' property value
Private m_StrsProxyAddress As String
' Variable to hold 'bUseProxyLogin' property value
Private m_bbUseProxyLogin  As Boolean
' Variable to hold 'sProxyUser' property value
Private m_StrsProxyUser    As String
' Variable to hold 'sProxyPass' property value
Private m_StrsProxyPass    As String
' Variable to hold 'AutomatiRedirection' property value
Private m_bAutomatiRedirection As Boolean
' Variable to hold 'UseLogin' property value
Private m_bUseLogin As Boolean
' Variable to hold 'UseName' property value
Private m_StrUseName As String
' Variable to hold 'UsePassword' property value
Private m_StrUsePassword As String
' Variable to hold 'SaveCookietoIE' property value
Private m_bSaveCookietoIE As Boolean
' Variable to hold 'TimeOut' property value
Private m_intTimeOut As Long
' Variable to hold 'ReturnBytArry' property value
Private m_BytReturnBytArry() As Byte
' Variable to hold flag if m_BytReturnBytArry is dimensioned
Private m_BytRead As Boolean
' Variable to hold 'NocookieAndcache' property value
Private m_bNocookieAndcache As Boolean
' Variable to hold 'StatusCode' property value
Private m_intStatusCode As Long

Private Sub Class_Initialize()
    'default values
    'inIDE = (App.LogMode = 0)
    m_StrRequestMethod = [cGET]
    m_StrRequestCookies = vbNullString
    m_StrRequestHeaders = vbNullString
    m_StrReferer = vbNullString
    m_StrsUserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:60.0) Gecko/20100101 Firefox/60.0"
    m_bbUseProxy = False
    m_bbUseProxyIE = False
    m_bbUseProxyLogin = False
    m_StrsProxyAddress = vbNullString '"proxy.ups.com:8080
    m_bAutomatiRedirection = True
    m_bUseLogin = False
    m_StrUsePassword = vbNullString
    m_StrUseName = vbNullString
    m_bSaveCookietoIE = False
    m_intTimeOut = 5000
    m_bNocookieAndcache = True
End Sub

' main method
Public Function API_Http(Optional sURL As String, _
                         Optional sMethod As Method, _
                         Optional sRequestHeads As String = "", _
                         Optional sCookie As String = "", _
                         Optional sPostData As String = "")

    On Error GoTo ErrorHandler:

    Dim iRetVal          As Long
    Dim sBuffer          As String * 1024
    Dim lBufferLen       As Long
    Dim vDllVersion      As tWinInetDLLVersion
    Dim sStatus          As String
    Dim sOptionBuffer    As String
    Dim lOptionBufferLen As Long
    Dim SecFlag          As Long
    Dim dwSecFlag        As Long
    Dim dwPort           As Long
    Dim hInternetConnect As Long, hHttpOpenRequest As Long
    Dim HttpORhttps      As String, sHost As String, sPort As String, sURI As String
    Dim bUseSecure       As Boolean, lService As Long
    Dim hInternetSession As Long
    Dim eCertAuthTry As Boolean, eProxyAuthTry As Boolean, eHttpAuthTry As Boolean
    Dim i As Long
    
    'substitute class level values to optional arguments
    If sURL = "" Then sURL = m_StrURL
    If sMethod = 0 Then sMethod = m_StrRequestMethod
    If sRequestHeads = "" Then sRequestHeads = m_StrRequestHeaders
    If sCookie = "" Then sCookie = m_StrRequestCookies
    If sPostData = "" Then sPostData = m_VarPostDatas
    Erase m_BytReturnBytArry
    
    Screen.MousePointer = vbHourglass
    lBufferLen = Len(sPostData)
    
    'see also INTERNET_OPTION_PER_CONNECTION_OPTION
    
    If m_bbUseProxyIE Then
        hInternetSession = InternetOpen(StrPtr(m_StrsUserAgent), INTERNET_OPEN_TYPE_PRECONFIG, 0&, 0&, 0&)
    ElseIf m_bbUseProxySocks4 Then
        hInternetSession = InternetOpen(StrPtr(m_StrsUserAgent), INTERNET_OPEN_TYPE_PROXY, StrPtr("SOCKS=" & m_StrsProxyAddress), 0&, 0&)
    ElseIf m_bbUseProxy Then
        hInternetSession = InternetOpen(StrPtr(m_StrsUserAgent), INTERNET_OPEN_TYPE_PROXY, StrPtr(m_StrsProxyAddress), 0&, 0&)
    Else
        hInternetSession = InternetOpen(StrPtr(m_StrsUserAgent), INTERNET_OPEN_TYPE_DIRECT, 0&, 0&, 0&)
    End If
    
    '// select which port should API use (cocus 19-feb-14)
    '// parse the url (cocus 19-feb-14)
    If Not ParseURL(sURL, HttpORhttps, m_StrUseName, m_StrUsePassword, sHost, sPort, sURI, , "-1") Then
        '// error parsing url
        Exit Function
    End If
    
    Select Case LCase$(HttpORhttps)

        Case "http":                lService = INTERNET_SERVICE_HTTP

        Case "https":               lService = INTERNET_SERVICE_HTTP: bUseSecure = True

        Case "ftp":                 lService = INTERNET_SERVICE_FTP
    End Select
    
    dwPort = CLng(sPort)

    If CBool(hInternetSession) Then
        '
        InternetQueryOption hInternetSession, INTERNET_OPTION_VERSION, VarPtr(vDllVersion), Len(vDllVersion)

        '        Debug.Print vDllVersion.lMajorVersion
        '        Debug.Print vDllVersion.lMinorVersion
        '        Debug.Print "InternetConnect"
        If bUseSecure = True Then
          
            SecFlag = INTERNET_FLAG_SECURE Or INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
        
        End If

        '// select which port should API use (cocus 19-feb-14)
        If dwPort = -1 Then
            If LCase$(sPort) = "ftp" Then
                dwPort = INTERNET_DEFAULT_FTP_PORT
            ElseIf bUseSecure Then
                dwPort = INTERNET_DEFAULT_HTTPS_PORT
            Else
                dwPort = INTERNET_DEFAULT_HTTP_PORT
            End If
        End If

        If m_bAutomatiRedirection = False Then
        
            SecFlag = INTERNET_FLAG_NO_AUTO_REDIRECT
        End If

        If m_bNocookieAndcache = True Then
            SecFlag = SecFlag Or (INTERNET_FLAG_NO_CACHE_WRITE + INTERNET_FLAG_NO_COOKIES)
        End If
        '
        hInternetConnect = InternetConnect(hInternetSession, StrPtr(sHost), dwPort, StrPtr(m_StrUseName), StrPtr(m_StrUsePassword), lService, 0, 0)

        If hInternetConnect > 0 Then
            'Debug.Print "HttpOpenRequest"

            If sMethod = 1 Then
                sOptionBuffer = vbNullString
                lOptionBufferLen = 0
                hHttpOpenRequest = HttpOpenRequest(hInternetConnect, StrPtr("GET"), StrPtr(sURI), StrPtr("HTTP/1.1"), 0, 0, INTERNET_FLAG_RELOAD Or SecFlag, 0)
            Else
                sOptionBuffer = sPostData
                lOptionBufferLen = Len(sPostData)
                hHttpOpenRequest = HttpOpenRequest(hInternetConnect, StrPtr("POST"), StrPtr(sURI), StrPtr("HTTP/1.1"), 0, 0, INTERNET_FLAG_RELOAD Or SecFlag, 0)
            End If

            If CBool(hHttpOpenRequest) Then
                'Debug.Print "HttpSendRequest"
                'Debug.Print sOptionBuffer
                
                Dim sHeader As String

                If InStr(sRequestHeads, "Accept:") = 0 Then
                    sRequestHeads = sRequestHeads + vbCrLf + "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*" + vbCrLf
                End If

                If InStr(sRequestHeads, "Referer:") = 0 Then
                    sRequestHeads = sRequestHeads + "Referer: " + sURL + vbCrLf
                End If

                sHeader = sRequestHeads
                sRequestHeads = vbNullString
                
                iRetVal = HttpAddRequestHeaders(hHttpOpenRequest, StrPtr(sHeader), Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)

                If sCookie <> "" Then
                    sCookie = "Cookie: " + sCookie + vbCrLf
                    ' szHeaders = szHeaders + sCookie
                    iRetVal = HttpAddRequestHeaders(hHttpOpenRequest, StrPtr(sCookie), Len(sCookie), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
                End If
                
                Dim dwTimeOut As Long
                dwTimeOut = m_intTimeOut                                    ' time out is set to 5 seconds
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_CONNECT_TIMEOUT, VarPtr(dwTimeOut), 4)
                'Debug.Print iRetVal & " " & Err.LastDllError & " " & "INTERNET_OPTION_CONNECT_TIMEOUT"
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, VarPtr(dwTimeOut), 4)
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_RECEIVE_TIMEOUT"
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SEND_TIMEOUT, VarPtr(dwTimeOut), 4)
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_SEND_TIMEOUT"
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_DATA_RECEIVE_TIMEOUT, VarPtr(dwTimeOut), 4)
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_DATA_RECEIVE_TIMEOUT"
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_DATA_SEND_TIMEOUT, VarPtr(dwTimeOut), 4)
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_DATA_SEND_TIMEOUT"
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_DISCONNECTED_TIMEOUT, VarPtr(dwTimeOut), 4) 'not suitable
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_DISCONNECTED_TIMEOUT"
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_FROM_CACHE_TIMEOUT, VarPtr(dwTimeOut), 4)
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_FROM_CACHE_TIMEOUT"
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_LISTEN_TIMEOUT, VarPtr(dwTimeOut), 4)
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_LISTEN_TIMEOUT"
                iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_REQFLAG_NET_TIMEOUT, VarPtr(dwTimeOut), 4)  'not suitable
                'Debug.Print iRetVal & " " & "INTERNET_REQFLAG_NET_TIMEOUT"
                
                If sPostData <> "" Then
                    If InStr(sHeader, "Content-Type:") = 0 Then
                        sRequestHeads = sRequestHeads + "Content-Type: application/x-www-form-urlencoded" + vbCrLf
                    End If
                    sRequestHeads = sRequestHeads + "Content-Length: " + CStr(Len(sPostData)) + vbCrLf
                End If
Resend:         '
                iRetVal = HttpSendRequest(hHttpOpenRequest, StrPtr(sRequestHeads), Len(sRequestHeads), StrPtr(sOptionBuffer), lOptionBufferLen)
                
                '
                ' Checking for errors
                '
                
                If (iRetVal <> 1) Then
                    If (Err.LastDllError = 12045) Then
                        If eCertAuthTry = 0 Then
                            '
                            'Certificate Authority is invalid.
                            'Debug.Print "Invalid Cert Auth, resending"
                            sStatus = "Invalid Cert Auth"
                            dwSecFlag = SECURITY_FLAG_IGNORE_UNKNOWN_CA
                            '.
                            iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SECURITY_FLAGS, VarPtr(dwSecFlag), 4)
                            'Debug.Print iRetVal & " " & Err.LastDllError & " " & "INTERNET_OPTION_SECURITY_FLAGS"
                            eCertAuthTry = 1
                            GoTo Resend
                        End If
                    Else
                        sStatus = ErrMessageText(Err.LastDllError)
                    End If
                Else
                    sStatus = ""
                End If
                
                If iRetVal Then
                    Dim dwStatus As Long, dwStatusSize As Long, dwIdxHeader As Long
                    dwIdxHeader = 0
                    dwStatusSize = LenB(dwStatus)
                    
                    If HttpQueryInfo(hHttpOpenRequest, HTTP_QUERY_FLAG_NUMBER Or HTTP_QUERY_STATUS_CODE, VarPtr(dwStatus), VarPtr(dwStatusSize), VarPtr(dwIdxHeader)) Then
                        
                        m_intStatusCode = dwStatus
                        
                        If dwStatus = HTTP_STATUS_PROXY_AUTH_REQ Then
                            sStatus = "Proxy authentification is required"
                            If m_bbUseProxyLogin = True Then
                                If eProxyAuthTry = 0 Then
                                    iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_PROXY_USERNAME, StrPtr(m_StrsProxyUser), LenB(m_StrsProxyUser) + 2)
                                    iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_PROXY_PASSWORD, StrPtr(m_StrsProxyPass), LenB(m_StrsProxyPass) + 2)
                                    eProxyAuthTry = 1
                                    GoTo Resend
                                End If
                            End If
                        ElseIf dwStatus = HTTP_STATUS_DENIED Then
                            sStatus = "Authentification is required"
                            If m_bUseLogin = True Then
                                If eHttpAuthTry = 0 Then
                                    iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_USERNAME, StrPtr(m_StrUseName), LenB(m_StrUseName) + 2)
                                    iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_PASSWORD, StrPtr(m_StrUsePassword), LenB(m_StrUsePassword) + 2)
                                    eHttpAuthTry = 1
                                    GoTo Resend
                                End If
                            End If
                        Else
                            sStatus = ""
                        End If
                    End If
                    
                    'Debug.Print "HttpQueryInfo"
                    'response headers ReturnHeads
                    Dim lblContentType    As String, lblContentLength As String, lblLastModified As String
                    Dim lblVersion        As String, lblStatusCode As String, lblRawHeaders As String, lblStatusText As String
                    Dim txtRequestHeaders As String
                    Dim lblForwarded      As String, lblServer As String, lblRequestMethod As String, lblPragma As String, lblUserAgent As String, lblRequestMethod2 As String
                    '                    GetQueryInfo hHttpOpenRequest, lblContentType, HTTP_QUERY_CONTENT_TYPE
                    '                    GetQueryInfo hHttpOpenRequest, lblContentLength, HTTP_QUERY_CONTENT_LENGTH
                    '                    GetQueryInfo hHttpOpenRequest, lblLastModified, HTTP_QUERY_LAST_MODIFIED
                    '                    GetQueryInfo hHttpOpenRequest, lblVersion, HTTP_QUERY_VERSION
                    '                    GetQueryInfo hHttpOpenRequest, lblStatusCode, HTTP_QUERY_STATUS_CODE
                    '                    GetQueryInfo hHttpOpenRequest, lblStatusText, HTTP_QUERY_STATUS_TEXT
                    '                    GetQueryInfo hHttpOpenRequest, lblRawHeaders, HTTP_QUERY_RAW_HEADERS
                    GetQueryInfo hHttpOpenRequest, m_StrRequestHeaders, HTTP_QUERY_RAW_HEADERS_CRLF
                    
                    '                    GetQueryInfo hHttpOpenRequest, lblForwarded, HTTP_QUERY_FORWARDED
                    '                    GetQueryInfo hHttpOpenRequest, lblServer, HTTP_QUERY_SERVER
                    '                    GetQueryInfo hHttpOpenRequest, lblRequestMethod, HTTP_QUERY_REQUEST_METHOD
                    '                    GetQueryInfo hHttpOpenRequest, lblPragma, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_PRAGMA
                    '                    GetQueryInfo hHttpOpenRequest, txtRequestHeaders, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_RAW_HEADERS_CRLF
                    '                    GetQueryInfo hHttpOpenRequest, lblUserAgent, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_USER_AGENT
                    '                    GetQueryInfo hHttpOpenRequest, lblRequestMethod2, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_REQUEST_METHOD
                    
                    If InStr(m_StrRequestHeaders, "Set-Cookie: ") > 0 Then
                        
                        m_StrRequestCookies = LiteCookies(m_StrRequestHeaders)

                        If m_bSaveCookietoIE = True Then
                            Dim buff() As String
                            buff = Split(m_StrRequestCookies, ";")

                            'InternetSetCookie "http://" + CheckUrl(sURL), "cookie:", ReturnCookies'
                            For i = 0 To UBound(buff) - 1
                                InternetSetCookie "http://" + CheckUrl(sURL), Split(buff(i), "=")(0), Split(buff(i), "=")(1) '& ";expires=Sun,22-Feb-2099 00:00:00 GMT;path=/; "
                            Next
                        End If
                    End If
                    
                    If sStatus = "" Then
                        If Len(m_StrRequestHeaders) > 0 Then
                            Dim aLine() As String
                            aLine = Split(m_StrRequestHeaders, vbCrLf)
                            For i = 0 To UBound(aLine)
                                If UCase$(Left(aLine(i), 7)) = "STATUS:" Then
                                    sStatus = Mid$(aLine(i), 8)
                                End If
                            Next
                            If Len(sStatus) = 0 Then
                                For i = 0 To UBound(aLine)
                                    If InStr(aLine(i), CStr(dwStatus)) <> 0 Then
                                        sStatus = aLine(i)
                                    End If
                                Next
                            End If
                        End If
                    End If
                    
                    If m_intStatusCode = 404 Then
                        API_Http = False
                        GoTo CloseHandles
                    End If
                    
                    If dwStatus <> HTTP_STATUS_PROXY_AUTH_REQ And dwStatus <> HTTP_STATUS_DENIED And dwStatus <> 400 Then
                        If sStatus = "" Then
                            sStatus = "ok~~~"
                        End If
                    Else
                        API_Http = False
                        GoTo CloseHandles
                    End If
                    
                    'RaiseEvent RequestHeads(sURL, m_StrRequestHeaders)
                    '
                    Dim lBytesRead As Long, bBuffer() As Byte, lRead As Long
                    ReDim bBuffer(0 To CHUNK_SIZE)
                    Do
                        InternetReadFile hHttpOpenRequest, bBuffer(0), CHUNK_SIZE, lRead

                        If lRead Then
                            m_BytRead = True
                            ReDim Preserve m_BytReturnBytArry(0 To lBytesRead + lRead - 1&)
                            memcpy m_BytReturnBytArry(lBytesRead), bBuffer(0), lRead
                            lBytesRead = lBytesRead + lRead
                        Else
                            Exit Do
                        End If
                    Loop
                    API_Http = True
                    'RaiseEvent RequestFinished(sURL, m_BytReturnBytArry, lBytesRead)
                Else
                    ' HttpSendRequest failed
                    sStatus = "HttpSendRequest call failed; Error code: " & Err.LastDllError & "."
                    API_Http = False
                    'RaiseEvent RequestError(sURL, sStatus)
                End If
            Else
                ' HttpOpenRequest failed
                sStatus = "HttpOpenRequest call failed; Error code: " & Err.LastDllError & "."
                API_Http = False
                
            End If
        Else
            ' InternetConnect failed
            sStatus = "InternetConnect call failed; Error code: " & Err.LastDllError & "."
            API_Http = False
        End If
    Else
        ' hInternetSession handle not allocated
        sStatus = "InternetOpen call failed: Error code: " & Err.LastDllError & "."
        API_Http = False
    End If
CloseHandles:
    m_ReturnStatus = sStatus
    If hHttpOpenRequest <> 0 Then InternetCloseHandle hHttpOpenRequest
    If hInternetConnect <> 0 Then InternetCloseHandle hInternetConnect
    If hInternetSession <> 0 Then InternetCloseHandle hInternetSession
    Screen.MousePointer = 0
    Exit Function
ErrorHandler:
    ErrorMsg Err, "API_Http", sURL, sMethod
    If inIDE Then Stop: Resume Next
End Function



Private Function CheckUrl(ByVal txtURL As String) As String
    On Error GoTo ErrorHandler:
    'Host: translate.google.cn
    If Right$(txtURL, 1) = "/" Then
        txtURL = Mid$(txtURL, 1, Len(txtURL) - 1)
    End If
    txtURL = Replace(txtURL, "http://", "")

    If Len(txtURL) = 0 Then txtURL = "www.microsoft.com"
    Dim posSlash As Long
    posSlash = InStr(txtURL, "/")

    If InStr(txtURL, "/") <> 0 Then
        CheckUrl = Left(txtURL, InStr(txtURL, "/") - 1)
    Else
        CheckUrl = txtURL
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "CheckUrl", txtURL
    If inIDE Then Stop: Resume Next
End Function

Private Function GetUrlObject(ByVal txtURL As String) As String
    On Error GoTo ErrorHandler:
    'GET /translate_a/t?client=t&hl=zh-CN&sl=en&tl=zh-CN&ie=UTF-8&oe=UTF-8&prev=btn&ssel=0&tsel=0&sc=1&q=Establishing%20secure%20connection HTTP/1.1
    If Right$(txtURL, 1) = "/" Then
        txtURL = Mid$(txtURL, 1, Len(txtURL) - 1)
    End If
    txtURL = Replace(txtURL, "http://", "")

    If InStr(txtURL, "/") <> 0 Then
        GetUrlObject = Right(txtURL, Len(txtURL) - InStr(txtURL, "/") + 1)
    Else
        GetUrlObject = ""
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetUrlObject", txtURL
    If inIDE Then Stop: Resume Next
End Function

Private Function GetQueryInfo(hHttpRequest As Long, s_outRequest As String, iInfoLevel As Long) As Boolean
    On Error GoTo ErrorHandler:
    Dim sBuffer       As String
    Dim lBufferLength As Long
    Dim dwIndex As Long
    Dim pos As Long
    dwIndex = 0
    sBuffer = String(5000, 0&)
    lBufferLength = LenB(sBuffer)
    GetQueryInfo = HttpQueryInfo(hHttpRequest, iInfoLevel, StrPtr(sBuffer), VarPtr(lBufferLength), VarPtr(dwIndex))
    If GetQueryInfo Then
        pos = InStr(sBuffer, vbNullChar)
        If pos <> 0 Then s_outRequest = Left$(sBuffer, pos - 1)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetQueryInfo", hHttpRequest, iInfoLevel
    If inIDE Then Stop: Resume Next
End Function

Public Function LiteCookies(cookies As String) As String
    On Error GoTo ErrorHandler:
    Dim rStr   As String
    'Dim oRegexp As Object
    Dim matchs As Object, match As Object
    cookies = Replace$(cookies, ";HttpOnly", ";", , , vbTextCompare)
    cookies = Replace$(cookies, vbCrLf, ";" & vbCrLf)
    cookies = Replace$(cookies, ";;", ";")
    cookies = Replace$(cookies, ";", "; ")

    'Set oRegexp = CreateObject("vbscript.regExp")
    
    If Not g_bRegexpInit Then
        Set oRegexp = New cRegExp
        g_bRegexpInit = True
    End If
    
    oRegexp.Global = True
    oRegexp.IgnoreCase = True
    oRegexp.MultiLine = True
    oRegexp.Pattern = "Set-Cookie: (.+?);"
    Set matchs = oRegexp.Execute(cookies)

    For Each match In matchs
        'Debug.Print match.Value
        ' Debug.Print match.SubMatches(0)
        rStr = GetCleanCookie(rStr, match.SubMatches(0) & ";")
    Next
    LiteCookies = rStr
    Set match = Nothing
    Set matchs = Nothing
    Set oRegexp = Nothing
    Exit Function
ErrorHandler:
    ErrorMsg Err, "LiteCookies", cookies
    If inIDE Then Stop: Resume Next
End Function

Private Function GetCleanCookie(Source As String, inStrs As String) As String
    On Error GoTo ErrorHandler:
    Dim Changed As Boolean, tem() As String
    Dim i       As Long

    If Source <> "" And inStrs <> "" Then
        Changed = False
        tem = Split(Source, ";")

        For i = 0 To UBound(tem) - 1
             
            If Split(tem(i), "=")(0) = Split(inStrs, "=")(0) Then

                If Split(inStrs, "=")(1) <> "deleted;" Then
                    Source = Replace(Source, tem(i) & ";", inStrs)
                    Changed = True
                End If
                GetCleanCookie = Source
                Exit Function
            End If
            
        Next

        If Changed = False And Split(inStrs, "=")(1) <> "deleted;" Then
            GetCleanCookie = Source + inStrs
        Else
            GetCleanCookie = Source
        End If
        
    ElseIf Source <> "" Then
        GetCleanCookie = Source
    ElseIf inStrs <> "" Then
        GetCleanCookie = inStrs
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetCleanCookie", Source, inStrs
    If inIDE Then Stop: Resume Next
End Function


Private Function ParseURL(ByVal sURL As String, _
                          ByRef sProtocol As String, _
                          ByRef sUser As String, _
                          ByRef sPassword As String, _
                          ByRef sHost As String, _
                          ByRef sPort As String, _
                          ByRef sURI As String, _
                          Optional ByVal sDefaultProtocol As String = "http", _
                          Optional ByVal sDefaultPort As String = "80") As Boolean
    
    On Error GoTo ErrorHandler:
    
    Dim iProtocolDelimiter As Long
    Dim iLoginDelimiter    As Long
    Dim iLoginPass         As Long
    Dim iHostDelimiter     As Long
    Dim iHostPort          As Long
    Dim sTemp              As String
    
    '// obtain the protocol delimiter
    iProtocolDelimiter = InStr(1, sURL, "://") - 1

    If iProtocolDelimiter = -1 Then
        sProtocol = sDefaultProtocol
    Else
        sProtocol = Left$(sURL, iProtocolDelimiter)
        
        '// trim out the protocol
        sURL = Right$(sURL, Len(sURL) - iProtocolDelimiter - 3)
    End If

    '// obtain the user & pass
    iLoginDelimiter = InStr(1, sURL, "@") - 1

    If Not (iLoginDelimiter = -1) Then
        sTemp = Left$(sURL, iLoginDelimiter)
        iLoginPass = InStr(1, sTemp, ":") - 1
        
        If Not (iLoginPass = -1) Then
            sUser = Left$(sTemp, iLoginPass)
            sPassword = Right$(sTemp, Len(sTemp) - iLoginPass - 1)
        Else
            sUser = sTemp
        End If
        
        '// trim out the login
        sURL = Right$(sURL, Len(sURL) - iLoginDelimiter - 1)
    End If
    
    '// obtain the host
    iHostDelimiter = InStr(1, sURL, "/") - 1

    If iHostDelimiter = -1 Then

        '// WTF?
        If Len(sURL) = 0 Then
            Exit Function
        End If
        
        sTemp = sURL
        sURL = ""
        sPort = sDefaultPort
    Else
        sTemp = Left$(sURL, iHostDelimiter)
        
        '// trim out the host
        sURL = Right$(sURL, Len(sURL) - iHostDelimiter)
    End If
    
    '// grab the host and its port
    iHostPort = InStr(1, sTemp, ":") - 1

    If iHostPort = -1 Then
        sHost = sTemp
        sPort = sDefaultPort
    Else
        sHost = Left$(sTemp, iHostPort)
        sPort = Right$(sTemp, Len(sTemp) - iHostPort - 1)
    End If

    If Len(sURL) = 0 Then
        sURI = "/"
    Else
        sURI = sURL
    End If
    
    ParseURL = True
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ParseURL", sURL
    If inIDE Then Stop: Resume Next
End Function

Private Function ErrMessageText(lCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
    Const FORMAT_MESSAGE_FROM_HMODULE As Long = 2048&
    Const MSG_SIZE = 300&
    
    Dim sRtrnMessage    As String
    Dim lret            As Long
    Dim hModule         As Long
    
    hModule = GetModuleHandle(StrPtr("wininet.dll"))
    
    sRtrnMessage = String$(MSG_SIZE, vbNullChar)
    lret = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, hModule, lCode, 0&, StrPtr(sRtrnMessage), MSG_SIZE, ByVal 0&)
    If lret > 0 Then
        ErrMessageText = Left$(sRtrnMessage, lret)
        ErrMessageText = Replace$(ErrMessageText, vbCrLf, vbNullString)
    End If
End Function



Public Property Let NocookieAndcache(ByVal bValue As Boolean)
    m_bNocookieAndcache = bValue
End Property

Public Property Get ReturnBytArry() As Byte()
    If Not m_BytRead Then Exit Property
    ReturnBytArry = m_BytReturnBytArry
End Property

Public Property Get ReturnString() As String
    If Not m_BytRead Then Exit Property
    ReturnString = String(UBound(m_BytReturnBytArry) \ 2 + 1, 0)
    memcpy ByVal StrPtr(ReturnString), m_BytReturnBytArry(0), UBound(m_BytReturnBytArry) + 1
    ReturnString = StrConv(ReturnString, vbUnicode)
    Dim pos As Long
    pos = InStr(ReturnString, vbNullChar)
    If pos > 0 Then ReturnString = Left(ReturnString, pos - 1)
End Property

Public Property Let TimeOut(ByVal intValue As Long)
    m_intTimeOut = intValue
End Property

Public Property Let SaveCookietoIE(ByVal bValue As Boolean)
    m_bSaveCookietoIE = bValue
End Property

Public Property Let UsePassword(ByVal StrValue As String)
    m_StrUsePassword = StrValue
End Property

Public Property Let UseName(ByVal StrValue As String)
    m_StrUseName = StrValue
End Property

Public Property Let UseAuthorization(ByVal bValue As Boolean)
    m_bUseLogin = bValue
End Property

Public Property Let AutomatiRedirection(ByVal bValue As Boolean)
    m_bAutomatiRedirection = bValue
End Property

Public Property Let ProxyPass(ByVal StrValue As String)
    m_StrsProxyPass = StrValue
End Property

Public Property Let ProxyUser(ByVal StrValue As String)
    m_StrsProxyUser = StrValue
End Property

Public Property Let UseProxyAuthorization(ByVal bValue As Boolean)
    m_bbUseProxyLogin = bValue
End Property
Public Property Get UseProxyAuthorization() As Boolean
    UseProxyAuthorization = m_bbUseProxyLogin
End Property

Public Property Let ProxyAddress(ByVal StrValue As String)
    m_StrsProxyAddress = StrValue
End Property

Public Property Let UseProxy(ByVal bValue As Boolean)
    m_bbUseProxy = bValue
End Property
Public Property Get UseProxy() As Boolean
    UseProxy = m_bbUseProxy
End Property
Public Property Let UseProxyIE(ByVal bValue As Boolean)
    m_bbUseProxyIE = bValue
End Property
Public Property Get UseProxyIE() As Boolean
    UseProxyIE = m_bbUseProxyIE
End Property
Public Property Let UseProxySocks4(ByVal bValue As Boolean)
    m_bbUseProxySocks4 = bValue
End Property
Public Property Get UseProxySocks4() As Boolean
    UseProxySocks4 = m_bbUseProxySocks4
End Property

Public Property Get UserAgent() As String
    UserAgent = m_StrsUserAgent
End Property
Public Property Let UserAgent(ByVal StrValue As String)
    m_StrsUserAgent = StrValue
End Property

Public Property Let Referer(ByVal StrValue As String)
    m_StrReferer = StrValue
End Property

Public Property Let PostDatas(ByVal VarValue As Variant)
    m_VarPostDatas = VarValue
End Property

Public Property Get URL() As String
    URL = m_StrURL
End Property
Public Property Let URL(ByVal StrValue As String)
    m_StrURL = StrValue
End Property

Public Property Get RequestMethod() As Method
    RequestMethod = m_StrRequestMethod
End Property
Public Property Let RequestMethod(ByVal StrValue As Method)
    m_StrRequestMethod = StrValue
End Property

Public Property Get RequestCookies() As String
    RequestCookies = m_StrRequestCookies
End Property
Public Property Let RequestCookies(ByVal StrValue As String)
    m_StrRequestCookies = StrValue
End Property

Public Property Get RequestHeaders() As String
    RequestHeaders = m_StrRequestHeaders
End Property
Public Property Let RequestHeaders(ByVal StrValue As String)
    m_StrRequestHeaders = StrValue
End Property

Public Property Get ReturnStatus() As String
    ReturnStatus = m_ReturnStatus
End Property

Public Property Get ReturnCode() As Long
    ReturnCode = m_intStatusCode
End Property
