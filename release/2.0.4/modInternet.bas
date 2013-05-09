Attribute VB_Name = "modInternet"
Option Explicit

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_OVERWRITEPROMPT = &H2

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private sUserAgent$
'Private Const sURLUpdate$ = "http://www.spywareinfo.com/~merijn/files/hijackthis-update.txt"
'Private Const sURLDownload$ = "http://www.spywareinfo.com/~merijn/files/hijackthis.zip"
Private Const sURLUpdate$ = vbNullString
Private Const sURLDownload$ = vbNullString

Public bDebug As Boolean
Public szResponse As String
Public szSubmitUrl As String


Public Sub CheckForUpdate()
    Dim hInternet&, hFile&, sBuffer$, lBufferLen&
    Dim sVer$, sUpdate$, sZipFile$, sThisVersion$
    Dim sProxy$, sFileName$
    'On Error GoTo Error:
    
    sThisVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    
    Dim szUpdateUrl As String
    szUpdateUrl = "http://www.trendmicro.com/go/hjt/check_for_updates/?hjtver=" & sThisVersion
        
    If True = IsOnline Then
        ShellExecute 0&, "open", szUpdateUrl, vbNullString, vbNullString, vbNormalFocus
    Else
        MsgBox "No Internet Connection Available"
    End If
    
    'sThisVersion = "1.97.4"
    
'    If MsgBox("HijackThis will now 'phone home' to www." & _
'       "spywareinfo.com and download a text" & _
'       "file with the latest version info.", vbInformation + vbOKCancel) = vbCancel Then
'        Exit Sub
'    End If
'
'    sUserAgent = "HijackThis v" & sThisVersion
'    sProxy = frmMain.txtCheckUpdateProxy.Text
'
'    hInternet = InternetOpen(sUserAgent, INTERNET_OPEN_TYPE_DIRECT, IIf(sProxy = vbNullString, vbNullString, sProxy), vbNullString, 0)
'    If hInternet = 0 Then
'        MsgBox "Unable to start WinSock.", vbCritical
'        Exit Sub
'    End If
'
'    hFile = InternetOpenUrl(hInternet, sURLUpdate, vbNullString, ByVal 0, INTERNET_FLAG_RELOAD, ByVal 0)
'    If hFile = 0 Then
'        MsgBox "Unable to connect to server (www.spywareinfo.com). Either you have no Internet connection open or the server is down.", vbCritical
'        InternetCloseHandle hInternet
'        Exit Sub
'    End If
'
'    sBuffer = Space(1024)
'    InternetReadFile hFile, sBuffer, 1024, lBufferLen
'    sBuffer = Left(sBuffer, lBufferLen)
'    InternetCloseHandle hFile
'
'    If Not IsNumeric(Left(sBuffer, 1)) Then
'        'oops. 404 or something?
'        MsgBox "The website is unavailable. Either it" & _
'               " is down, or " & _
'               "I moved or deleted it by accident. Try again" & _
'               " in a few days, or email me at www.merijn.org/contact.html.", vbCritical
'        InternetCloseHandle hInternet
'        Exit Sub
'    End If
'
'    sVer = Left(sBuffer, InStr(sBuffer, "|") - 1)
'    sUpdate = Mid(sBuffer, InStr(sBuffer, "|") + 1)
'
'    If CLng(Replace(sVer, ".", "")) > CLng(Replace(sThisVersion, ".", "")) Then
'        If MsgBox("You have: HijackThis v" & sThisVersion & vbCrLf & _
'               "The latest version is: HijackThis v" & sVer & vbCrLf & vbCrLf & _
'               "New in " & sVer & ": " & vbCrLf & sUpdate & vbCrLf & vbCrLf & _
'               "Do you want to download the new version?", vbQuestion + vbYesNo) = vbNo Then
'            Exit Sub
'        End If
'    Else
'        MsgBox "You have the latest version of HijackThis: " & sVer, vbInformation
'        InternetCloseHandle hInternet
'        Exit Sub
'    End If
'
'    hFile = InternetOpenUrl(hInternet, sURLDownload, vbNullString, ByVal 0, INTERNET_FLAG_RELOAD, ByVal 0)
'    If hFile = 0 Then
'        MsgBox "Unable to connect to server (www.spywareinfo.com). Either you have no Internet connection open or the server is down.", vbCritical
'        InternetCloseHandle hInternet
'        Exit Sub
'    End If
'
'    'improved sub for downloading hijackthis.zip
'    'it might not come in into one piece
'    sZipFile = vbNullString
'    Do
'        sBuffer = Space(32768)
'        InternetReadFile hFile, sBuffer, Len(sBuffer), lBufferLen
'        sZipFile = sZipFile & Left(sBuffer, lBufferLen)
'    Loop Until lBufferLen = 0
'
'    InternetCloseHandle hFile
'    InternetCloseHandle hInternet
'
'    sFileName$ = CmnDlgSaveFile("Save downloaded file to...", "All files (*.*)|*.*", "hijackthis.zip")
'    If sFileName = vbNullString Then
'        Exit Sub
'    End If
'
'    Open sFileName For Output As #1
'        Print #1, sZipFile
'    Close #1
'
'    If MsgBox("Newest version of HijackThis downloaded to" & vbCrLf & sFileName & vbCrLf & vbCrLf & "Open the file?", vbQuestion + vbYesNo) = vbNo Then
'        Exit Sub
'    End If
'
'    hFile = ShellExecute(0, "open", sFileName, vbNullString, vbNullString, 1)
'
'EndOfSub:
'    Exit Sub
'
'Error:
'    Close #1
'    ErrorMsg "modInternet_CheckForUpdate", Err.Number, Err.Description
End Sub

Public Sub SendData(szUrl As String, szData As String)
On Error GoTo Error
Dim szRequest As String
Dim xmlhttp As Object
Dim dataLen As Integer
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

szRequest = "data=" & URLEncode(szData)

dataLen = Len(szRequest)
xmlhttp.Open "POST", szUrl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'xmlhttp.setRequestHeader "User-Agent", "HJT.1.99.2" & "|" & sWinVersion & "|" & sMSIEVersion

xmlhttp.send "" & szRequest
'MsgBox szData

szResponse = xmlhttp.responseText
'MsgBox szResponse

Set xmlhttp = Nothing

Error:

End Sub

Function GetUrl(szUrl As String) As String
On Error GoTo Error:
Dim szRequest As String
Dim xmlhttp As Object
Dim dataLen As Integer
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

dataLen = Len(szRequest)
xmlhttp.Open "GET", szUrl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'xmlhttp.setRequestHeader "User-Agent", "HJT.1.99.2" & "|" & sWinVersion & "|" & sMSIEVersion

xmlhttp.send "" & szRequest
'MsgBox szData

GetUrl = xmlhttp.responseText
'MsgBox szResponse

Set xmlhttp = Nothing
Exit Function

Error:
GetUrl = "HJT_NOT_SUPPORTED"
End Function

Public Sub ParseHTTPResponse(szResponse As String)

Dim curPos As Integer
Dim startIDPos, endIDPos, startDataPos, endDataPos As Integer
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
    
    szDataId = Mid(szResponse, startIDPos, endIDPos - startIDPos)
    szData = Mid(szResponse, startDataPos, endDataPos - startDataPos)
    
    Select Case szDataId
    Case "REPORT_URL"
    ShellExecute 0&, "open", szData, vbNullString, vbNullString, vbNormalFocus
    Case "SUBMIT_URL"
    szSubmitUrl = szData
    End Select
    
Loop


End Sub
Function URLEncode(ByVal Text As String) As String
    Dim i As Integer
    Dim acode As Integer
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
    
End Function

Public Function IsOnline() As Boolean

   IsOnline = InternetGetConnectedState(0&, 0&)
     
End Function
