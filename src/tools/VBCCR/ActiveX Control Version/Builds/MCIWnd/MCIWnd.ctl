VERSION 5.00
Begin VB.UserControl MCIWnd 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "MCIWnd.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "MCIWnd.ctx":0026
End
Attribute VB_Name = "MCIWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private MciFormatMilliseconds, MciFormatHms, MciFormatMsf, MciFormatFrames, MciFormatSmpte24, MciFormatSmpte25, MciFormatSmpte30, MciFormatSmpte30Drop, MciFormatBytes, MciFormatSamples, MciFormatTmsf
Private MciModeNotReady, MciModeStop, MciModePlay, MciModeRecord, MciModeSeek, MciModePause, MciModeOpen
Private MciNotifySuccessful, MciNotifySuperseded, MciNotifyAborted, MciNotifyFailure
Private MciCaptionNone, MciCaptionName, MciCaptionNamePos, MciCaptionNameMode, MciCaptionNamePosMode, MciCaptionPos, MciCaptionPosMode, MciCaptionMode
#End If
Private Const MCI_FORMAT_MILLISECONDS As Long = 0
Private Const MCI_FORMAT_HMS As Long = 1
Private Const MCI_FORMAT_MSF As Long = 2
Private Const MCI_FORMAT_FRAMES As Long = 3
Private Const MCI_FORMAT_SMPTE_24 As Long = 4
Private Const MCI_FORMAT_SMPTE_25 As Long = 5
Private Const MCI_FORMAT_SMPTE_30 As Long = 6
Private Const MCI_FORMAT_SMPTE_30DROP As Long = 7
Private Const MCI_FORMAT_BYTES As Long = 8
Private Const MCI_FORMAT_SAMPLES As Long = 9
Private Const MCI_FORMAT_TMSF As Long = 10
Public Enum MciFormatConstants
MciFormatMilliseconds = MCI_FORMAT_MILLISECONDS
MciFormatHms = MCI_FORMAT_HMS
MciFormatMsf = MCI_FORMAT_MSF
MciFormatFrames = MCI_FORMAT_FRAMES
MciFormatSmpte24 = MCI_FORMAT_SMPTE_24
MciFormatSmpte25 = MCI_FORMAT_SMPTE_25
MciFormatSmpte30 = MCI_FORMAT_SMPTE_30
MciFormatSmpte30Drop = MCI_FORMAT_SMPTE_30DROP
MciFormatBytes = MCI_FORMAT_BYTES
MciFormatSamples = MCI_FORMAT_SAMPLES
MciFormatTmsf = MCI_FORMAT_TMSF
End Enum
Private Const MCI_STRING_OFFSET As Long = 512
Private Const MCI_MODE_NOT_READY As Long = (MCI_STRING_OFFSET + 12)
Private Const MCI_MODE_STOP As Long = (MCI_STRING_OFFSET + 13)
Private Const MCI_MODE_PLAY As Long = (MCI_STRING_OFFSET + 14)
Private Const MCI_MODE_RECORD As Long = (MCI_STRING_OFFSET + 15)
Private Const MCI_MODE_SEEK As Long = (MCI_STRING_OFFSET + 16)
Private Const MCI_MODE_PAUSE As Long = (MCI_STRING_OFFSET + 17)
Private Const MCI_MODE_OPEN As Long = (MCI_STRING_OFFSET + 18)
Public Enum MciModeConstants
MciModeNotReady = MCI_MODE_NOT_READY
MciModeStop = MCI_MODE_STOP
MciModePlay = MCI_MODE_PLAY
MciModeRecord = MCI_MODE_RECORD
MciModeSeek = MCI_MODE_SEEK
MciModePause = MCI_MODE_PAUSE
MciModeOpen = MCI_MODE_OPEN
End Enum
Private Const MCI_NOTIFY_SUCCESSFUL = &H1
Private Const MCI_NOTIFY_SUPERSEDED = &H2
Private Const MCI_NOTIFY_ABORTED = &H4
Private Const MCI_NOTIFY_FAILURE = &H8
Public Enum MciNotifyConstants
MciNotifySuccessful = MCI_NOTIFY_SUCCESSFUL
MciNotifySuperseded = MCI_NOTIFY_SUPERSEDED
MciNotifyAborted = MCI_NOTIFY_ABORTED
MciNotifyFailure = MCI_NOTIFY_FAILURE
End Enum
Public Enum MciCaptionConstants
MciCaptionNone = 0
MciCaptionName = 1
MciCaptionNamePos = 2
MciCaptionNameMode = 3
MciCaptionNamePosMode = 4
MciCaptionPos = 5
MciCaptionPosMode = 6
MciCaptionMode = 7
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Public Event ModeChange(ByVal NewMode As MciModeConstants)
Attribute ModeChange.VB_Description = "Occurs when the operating mode of the MCI device has changed."
Public Event PositionChange(ByVal NewPosition As Long)
Attribute PositionChange.VB_Description = "Occurs when the position has changed."
Public Event MediaChange(ByVal NewFileName As String)
Attribute MediaChange.VB_Description = "Occurs when the media has changed."
Public Event Error(ByVal ErrorCode As Long)
Public Event Notify(ByVal NotifyCode As MciNotifyConstants)
Attribute Notify.VB_Description = "Notification that an MCI device has completed an operation."
Public Event Signal(ByVal DeviceID As Long, ByVal dwUserParam As Long)
Attribute Signal.VB_Description = "Notification that an MCI device has reached a position defined in a previous signal command."
Public Event PreviewKeyDown(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyDown.VB_Description = "Occurs before the KeyDown event."
Public Event PreviewKeyUp(ByVal KeyCode As Integer, ByRef IsInputKey As Boolean)
Attribute PreviewKeyUp.VB_Description = "Occurs before the KeyUp event."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event KeyPress(KeyChar As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an character key."
Attribute KeyPress.VB_UserMemId = -603
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the user moves the mouse into the control."
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the user moves the mouse out of the control."
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MCIWndRegisterClass Lib "msvfw32" () As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringW" (ByVal lpAppName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const GW_CHILD As Long = 5
Private Const GW_HWNDNEXT As Long = 2
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CAPTION As Long = &HC00000
Private Const WM_CLOSE As Long = &H10
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_UNICHAR As Long = &H109, UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_SYSCOMMAND As Long = &H112, SC_MOVE As Long = &HF010&
Private Const MCIWNDF_NOAUTOSIZEWINDOW As Long = &H1
Private Const MCIWNDF_NOPLAYBAR As Long = &H2
Private Const MCIWNDF_NOAUTOSIZEMOVIE As Long = &H4
Private Const MCIWNDF_NOMENU As Long = &H8
Private Const MCIWNDF_SHOWNAME As Long = &H10
Private Const MCIWNDF_SHOWPOS As Long = &H20
Private Const MCIWNDF_SHOWMODE As Long = &H40
Private Const MCIWNDF_SHOWALL As Long = &H70
Private Const MCIWNDF_NOTIFYANSI As Long = &H80
Private Const MCIWNDF_NOTIFYMODE As Long = &H100
Private Const MCIWNDF_NOTIFYPOS As Long = &H200
Private Const MCIWNDF_NOTIFYSIZE As Long = &H400
Private Const MCIWNDF_NOTIFYMEDIA = &H880
Private Const MCIWNDF_NOTIFYERROR As Long = &H1000
Private Const MCIWNDF_NOTIFYALL As Long = &H1F00
Private Const MCIWNDF_RECORD As Long = &H2000
Private Const MCIWNDF_NOERRORDLG As Long = &H4000
Private Const MCIWNDF_NOOPEN As Long = &H8000&
Private Const WM_USER As Long = &H400
Private Const MCIWNDM_GETDEVICEID As Long = (WM_USER + 100)
Private Const MCIWNDM_SENDSTRINGA As Long = (WM_USER + 101)
Private Const MCIWNDM_SENDSTRINGW As Long = (WM_USER + 201)
Private Const MCIWNDM_SENDSTRING As Long = MCIWNDM_SENDSTRINGW
Private Const MCIWNDM_GETPOSITIONA As Long = (WM_USER + 102)
Private Const MCIWNDM_GETPOSITIONW As Long = (WM_USER + 202)
Private Const MCIWNDM_GETPOSITION As Long = MCIWNDM_GETPOSITIONW
Private Const MCIWNDM_GETSTART As Long = (WM_USER + 103)
Private Const MCIWNDM_GETLENGTH As Long = (WM_USER + 104)
Private Const MCIWNDM_GETEND As Long = (WM_USER + 105)
Private Const MCIWNDM_GETMODEA As Long = (WM_USER + 106)
Private Const MCIWNDM_GETMODEW As Long = (WM_USER + 206)
Private Const MCIWNDM_GETMODE As Long = MCIWNDM_GETMODEW
Private Const MCIWNDM_EJECT As Long = (WM_USER + 107)
Private Const MCIWNDM_SETZOOM As Long = (WM_USER + 108)
Private Const MCIWNDM_GETZOOM As Long = (WM_USER + 109)
Private Const MCIWNDM_SETVOLUME As Long = (WM_USER + 110)
Private Const MCIWNDM_GETVOLUME As Long = (WM_USER + 111)
Private Const MCIWNDM_SETSPEED As Long = (WM_USER + 112)
Private Const MCIWNDM_GETSPEED As Long = (WM_USER + 113)
Private Const MCIWNDM_SETREPEAT As Long = (WM_USER + 114)
Private Const MCIWNDM_GETREPEAT As Long = (WM_USER + 115)
Private Const MCIWNDM_SETTIMEFORMATA As Long = (WM_USER + 119)
Private Const MCIWNDM_SETTIMEFORMATW As Long = (WM_USER + 219)
Private Const MCIWNDM_SETTIMEFORMAT As Long = MCIWNDM_SETTIMEFORMATW
Private Const MCIWNDM_GETTIMEFORMATA As Long = (WM_USER + 120)
Private Const MCIWNDM_GETTIMEFORMATW As Long = (WM_USER + 220)
Private Const MCIWNDM_GETTIMEFORMAT As Long = MCIWNDM_GETTIMEFORMATW
Private Const MCIWNDM_VALIDATEMEDIA As Long = (WM_USER + 121)
Private Const MCIWNDM_PLAYFROM As Long = (WM_USER + 122)
Private Const MCIWNDM_PLAYTO As Long = (WM_USER + 123)
Private Const MCIWNDM_GETFILENAMEA As Long = (WM_USER + 124)
Private Const MCIWNDM_GETFILENAMEW As Long = (WM_USER + 224)
Private Const MCIWNDM_GETFILENAME As Long = MCIWNDM_GETFILENAMEW
Private Const MCIWNDM_GETDEVICEA As Long = (WM_USER + 125)
Private Const MCIWNDM_GETDEVICEW As Long = (WM_USER + 225)
Private Const MCIWNDM_GETDEVICE As Long = MCIWNDM_GETDEVICEW
Private Const MCIWNDM_GETERRORA As Long = (WM_USER + 128)
Private Const MCIWNDM_GETERRORW As Long = (WM_USER + 228)
Private Const MCIWNDM_GETERROR As Long = MCIWNDM_GETERRORW
Private Const MCIWNDM_SETTIMERS As Long = (WM_USER + 129)
Private Const MCIWNDM_SETACTIVETIMER As Long = (WM_USER + 130)
Private Const MCIWNDM_SETINACTIVETIMER As Long = (WM_USER + 131)
Private Const MCIWNDM_GETACTIVETIMER As Long = (WM_USER + 132)
Private Const MCIWNDM_GETINACTIVETIMER As Long = (WM_USER + 133)
Private Const MCIWNDM_NEWA As Long = (WM_USER + 134)
Private Const MCIWNDM_NEWW As Long = (WM_USER + 234)
Private Const MCIWNDM_NEW As Long = MCIWNDM_NEWW
Private Const MCIWNDM_CHANGESTYLES As Long = (WM_USER + 135)
Private Const MCIWNDM_GETSTYLES As Long = (WM_USER + 136)
Private Const MCIWNDM_GETALIAS As Long = (WM_USER + 137)
Private Const MCIWNDM_RETURNSTRINGA As Long = (WM_USER + 138)
Private Const MCIWNDM_RETURNSTRINGW As Long = (WM_USER + 238)
Private Const MCIWNDM_RETURNSTRING As Long = MCIWNDM_RETURNSTRINGW
Private Const MCIWNDM_PLAYREVERSE As Long = (WM_USER + 139)
Private Const MCIWNDM_GET_SOURCE As Long = (WM_USER + 140)
Private Const MCIWNDM_GET_DEST As Long = (WM_USER + 142)
Private Const MCIWNDM_PUT_DEST As Long = (WM_USER + 143)
Private Const MCIWNDM_CAN_PLAY As Long = (WM_USER + 144)
Private Const MCIWNDM_CAN_WINDOW As Long = (WM_USER + 145)
Private Const MCIWNDM_CAN_RECORD As Long = (WM_USER + 146)
Private Const MCIWNDM_CAN_SAVE As Long = (WM_USER + 147)
Private Const MCIWNDM_CAN_EJECT As Long = (WM_USER + 148)
Private Const MCIWNDM_CAN_CONFIG As Long = (WM_USER + 149)
Private Const MCIWNDM_OPENA As Long = (WM_USER + 153)
Private Const MCIWNDM_OPENW As Long = (WM_USER + 252)
Private Const MCIWNDM_OPEN As Long = MCIWNDM_OPENW
Private Const MCIWNDM_NOTIFYMODE As Long = (WM_USER + 200)
Private Const MCIWNDM_NOTIFYPOS As Long = (WM_USER + 201)
Private Const MCIWNDM_NOTIFYSIZE As Long = (WM_USER + 202)
Private Const MCIWNDM_NOTIFYMEDIA As Long = (WM_USER + 203)
Private Const MCIWNDM_NOTIFYERROR As Long = (WM_USER + 205)
Private Const MCI_SAVE As Long = &H813
Private Const MM_MCINOTIFY As Long = &H3B9
Private Const MM_MCISIGNAL As Long = &H3CB
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private MCIWndHandle As Long
Private MCIWndBackColorBrush As Long
Private MCIWndCharCodeCache As Long
Private MCIWndCommand As String
Private MCIWndIsClick As Boolean
Private MCIWndMouseOver As Boolean
Private MCIWndDesignMode As Boolean
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropBackColor As OLE_COLOR
Private PropBorderStyle As CCBorderStyleConstants
Private PropRepeat As Boolean
Private PropErrorDlg As Boolean
Private PropRecord As Boolean
Private PropPlaybar As Boolean
Private PropMenu As Boolean
Private PropAllowOpen As Boolean
Private PropAutoSizeWindow As Boolean
Private PropAutoSizeMovie As Boolean
Private PropTimerFreq As Integer
Private PropZoom As Long
Private PropCaption As MciCaptionConstants

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = wParam And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
            SendMessage hWnd, wMsg, wParam, ByVal lParam
            Handled = True
        Case vbKeyTab, vbKeyReturn, vbKeyEscape
            If IsInputKey = True Then
                SendMessage hWnd, wMsg, wParam, ByVal lParam
                Handled = True
            End If
    End Select
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetDisplayStringMousePointer(PropMousePointer, DisplayName)
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Then
    Value = Cookie
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
MCIWndRegisterClass
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
MCIWndDesignMode = Not Ambient.UserMode
On Error GoTo 0
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropBackColor = vbWindowBackground
PropBorderStyle = CCBorderStyleSingle
PropRepeat = False
PropErrorDlg = False
PropRecord = False
PropPlaybar = True
PropMenu = True
PropAllowOpen = True
PropAutoSizeWindow = True
PropAutoSizeMovie = True
PropTimerFreq = 500
PropZoom = 100
PropCaption = MciCaptionNone
Call CreateMCIWnd
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
MCIWndDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSingle)
PropRepeat = .ReadProperty("Repeat", False)
PropErrorDlg = .ReadProperty("ErrorDlg", False)
PropRecord = .ReadProperty("Record", False)
PropPlaybar = .ReadProperty("Playbar", True)
PropMenu = .ReadProperty("Menu", True)
PropAllowOpen = .ReadProperty("AllowOpen", True)
PropAutoSizeWindow = .ReadProperty("AutoSizeWindow", True)
PropAutoSizeMovie = .ReadProperty("AutoSizeMovie", True)
PropTimerFreq = .ReadProperty("TimerFreq", 500)
PropZoom = .ReadProperty("Zoom", 100)
PropCaption = .ReadProperty("Caption", MciCaptionNone)
End With
Call CreateMCIWnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSingle
.WriteProperty "Repeat", PropRepeat, False
.WriteProperty "ErrorDlg", PropErrorDlg, False
.WriteProperty "Record", PropRecord, False
.WriteProperty "Playbar", PropPlaybar, True
.WriteProperty "Menu", PropMenu, True
.WriteProperty "AllowOpen", PropAllowOpen, True
.WriteProperty "AutoSizeWindow", PropAutoSizeWindow, True
.WriteProperty "AutoSizeMovie", PropAutoSizeMovie, True
.WriteProperty "TimerFreq", PropTimerFreq, 500
.WriteProperty "Zoom", PropZoom, 100
.WriteProperty "Caption", PropCaption, MciCaptionNone
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
Static PrevHeight As Long, PrevWidth As Long
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If MCIWndHandle <> 0 Then MoveWindow MCIWndHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
InProc = False
If PrevHeight <> .ScaleHeight Or PrevWidth <> .ScaleWidth Then
    PrevHeight = .ScaleHeight
    PrevWidth = .ScaleWidth
    RaiseEvent Resize
End If
End With
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyMCIWnd
Call ComCtlsReleaseShellMod
End Sub

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name used in code to identify an object."
Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
Extender.Tag = Value
End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Returns the object on which this object is located."
Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
Attribute Container.VB_Description = "Returns the container of an object."
Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
Set Extender.Container = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns/sets the distance between the internal left edge of an object and the left edge of its container."
Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
Extender.Left = Value
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns/sets the distance between the internal top edge of an object and the top edge of its container."
Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
Extender.Top = Value
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of an object."
Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
Extender.Width = Value
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns/sets the height of an object."
Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
Extender.Height = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether an object is visible or hidden."
Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
Extender.Visible = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
Attribute ToolTipText.VB_MemberFlags = "400"
ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
Extender.ToolTipText = Value
End Property

Public Property Get HelpContextID() As Long
Attribute HelpContextID.VB_Description = "Specifies the default Help file context ID for an object."
HelpContextID = Extender.HelpContextID
End Property

Public Property Let HelpContextID(ByVal Value As Long)
Extender.HelpContextID = Value
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
Attribute WhatsThisHelpID.VB_MemberFlags = "400"
WhatsThisHelpID = Extender.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal Value As Long)
Extender.WhatsThisHelpID = Value
End Property

Public Property Get DragIcon() As IPictureDisp
Attribute DragIcon.VB_Description = "Returns/sets the icon to be displayed as the pointer in a drag-and-drop operation."
Attribute DragIcon.VB_MemberFlags = "400"
Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
Attribute DragMode.VB_Description = "Returns/sets a value that determines whether manual or automatic drag mode is used."
Attribute DragMode.VB_MemberFlags = "400"
DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
Extender.DragMode = Value
End Property

Public Sub Drag(Optional ByRef Action As Variant)
Attribute Drag.VB_Description = "Begins, ends, or cancels a drag operation of any object except Line, Menu, Shape, and Timer."
If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Sub SetFocus()
Attribute SetFocus.VB_Description = "Moves the focus to the specified object."
Extender.SetFocus
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
Attribute ZOrder.VB_Description = "Places a specified object at the front or back of the z-order within its graphical level."
If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
hWnd = MCIWndHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If MCIWndHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles MCIWndHandle
    Else
        RemoveVisualStyles MCIWndHandle
    End If
    Dim hWnd As Long
    hWnd = GetWindow(MCIWndHandle, GW_CHILD)
    Do While hWnd <> 0
        If PropVisualStyles = True Then
            ActivateVisualStyles hWnd
        Else
            RemoveVisualStyles hWnd
        End If
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If MCIWndHandle <> 0 Then
    EnableWindow MCIWndHandle, IIf(Value = True, 1, 0)
    Dim hWnd As Long
    hWnd = GetWindow(MCIWndHandle, GW_CHILD)
    Do While hWnd <> 0
        EnableWindow hWnd, IIf(Value = True, 1, 0)
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
End If
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDropMode() As OLEDropModeConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As OLEDropModeConstants)
Select Case Value
    Case OLEDropModeNone, OLEDropModeManual
        UserControl.OLEDropMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDropMode"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As Integer)
Select Case Value
    Case 0 To 16, 99
        PropMousePointer = Value
    Case Else
        Err.Raise 380
End Select
If MCIWndDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
Set MouseIcon = PropMouseIcon
End Property

Public Property Let MouseIcon(ByVal Value As IPictureDisp)
Set Me.MouseIcon = Value
End Property

Public Property Set MouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropMouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
        Set PropMouseIcon = Value
    Else
        If MCIWndDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If MCIWndDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get MouseTrack() As Boolean
Attribute MouseTrack.VB_Description = "Returns/sets whether mouse events occurs when the mouse pointer enters or leaves the control."
MouseTrack = PropMouseTrack
End Property

Public Property Let MouseTrack(ByVal Value As Boolean)
PropMouseTrack = Value
UserControl.PropertyChanged "MouseTrack"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If MCIWndHandle <> 0 Then
    If MCIWndBackColorBrush <> 0 Then DeleteObject MCIWndBackColorBrush
    MCIWndBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
End If
Me.Refresh
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As CCBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style."
Attribute BorderStyle.VB_UserMemId = -504
BorderStyle = PropBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As CCBorderStyleConstants)
Select Case Value
    Case CCBorderStyleNone, CCBorderStyleSingle, CCBorderStyleThin, CCBorderStyleSunken, CCBorderStyleRaised
        PropBorderStyle = Value
    Case Else
        Err.Raise 380
End Select
If MCIWndHandle <> 0 And PropCaption = MciCaptionNone Then Call ComCtlsChangeBorderStyle(MCIWndHandle, PropBorderStyle)
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get Repeat() As Boolean
Attribute Repeat.VB_Description = "Returns/sets a value indicating if the playback continuously repeats when the play button on the play bar is pressed."
If MCIWndHandle <> 0 Then
    Repeat = CBool(SendMessage(MCIWndHandle, MCIWNDM_GETREPEAT, 0, ByVal 0&) <> 0)
Else
    Repeat = PropRepeat
End If
End Property

Public Property Let Repeat(ByVal Value As Boolean)
PropRepeat = Value
If MCIWndHandle <> 0 Then SendMessage MCIWndHandle, MCIWNDM_SETREPEAT, 0, ByVal CLng(IIf(PropRepeat = True, 1, 0))
UserControl.PropertyChanged "Repeat"
End Property

Public Property Get ErrorDlg() As Boolean
Attribute ErrorDlg.VB_Description = "Returns/sets a value that determines whether MCI errors cause an error dialog to be displayed or not."
ErrorDlg = PropErrorDlg
End Property

Public Property Let ErrorDlg(ByVal Value As Boolean)
PropErrorDlg = Value
If MCIWndHandle <> 0 Then
    If PropErrorDlg = False Then
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOERRORDLG, ByVal MCIWNDF_NOERRORDLG
    Else
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOERRORDLG, ByVal 0&
    End If
End If
UserControl.PropertyChanged "ErrorDlg"
End Property

Public Property Get Record() As Boolean
Attribute Record.VB_Description = "Returns/sets a value indicating if recording controls or recording entries appear on the play bar or in the menus."
Record = PropRecord
End Property

Public Property Let Record(ByVal Value As Boolean)
PropRecord = Value
If MCIWndHandle <> 0 Then
    If PropRecord = True Then
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_RECORD, ByVal MCIWNDF_RECORD
    Else
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_RECORD, ByVal 0&
    End If
    Call UserControl_Resize
End If
UserControl.PropertyChanged "Record"
End Property

Public Property Get Playbar() As Boolean
Attribute Playbar.VB_Description = "Returns/sets a value indicating if a play bar appears in the control. The play bar lets the user control playback and recording of MCI devices."
Playbar = PropPlaybar
End Property

Public Property Let Playbar(ByVal Value As Boolean)
PropPlaybar = Value
If MCIWndHandle <> 0 Then
    If PropPlaybar = False Then
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOPLAYBAR, ByVal MCIWNDF_NOPLAYBAR
    Else
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOPLAYBAR, ByVal 0&
    End If
    Call UserControl_Resize
    Me.Enabled = UserControl.Enabled
End If
UserControl.PropertyChanged "Playbar"
End Property

Public Property Get Menu() As Boolean
Attribute Menu.VB_Description = "Returns/sets a value indicating if a menu appears on the play bar and if a right mouse-click over the control displays a pop-up menu."
Menu = PropMenu
End Property

Public Property Let Menu(ByVal Value As Boolean)
PropMenu = Value
If MCIWndHandle <> 0 Then
    If PropMenu = False Then
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOMENU, ByVal MCIWNDF_NOMENU
    Else
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOMENU, ByVal 0&
    End If
    Call UserControl_Resize
End If
UserControl.PropertyChanged "Menu"
End Property

Public Property Get AllowOpen() As Boolean
Attribute AllowOpen.VB_Description = "Returns/sets a value indicating if the control allows or prohibits users from accessing the open and close commands."
AllowOpen = PropAllowOpen
End Property

Public Property Let AllowOpen(ByVal Value As Boolean)
PropAllowOpen = Value
If MCIWndHandle <> 0 And MCIWndDesignMode = False Then
    If PropAllowOpen = False Then
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOOPEN, ByVal MCIWNDF_NOOPEN
    Else
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOOPEN, ByVal 0&
    End If
    Call UserControl_Resize
End If
UserControl.PropertyChanged "AllowOpen"
End Property

Public Property Get AutoSizeWindow() As Boolean
Attribute AutoSizeWindow.VB_Description = "Returns/sets a value that determines whether or not the control will change the dimensions of the window when the image size changes."
AutoSizeWindow = PropAutoSizeWindow
End Property

Public Property Let AutoSizeWindow(ByVal Value As Boolean)
PropAutoSizeWindow = Value
If MCIWndHandle <> 0 Then
    If PropAutoSizeWindow = False Then
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOAUTOSIZEWINDOW, ByVal MCIWNDF_NOAUTOSIZEWINDOW
    Else
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOAUTOSIZEWINDOW, ByVal 0&
    End If
    Call UserControl_Resize
End If
UserControl.PropertyChanged "AutoSizeWindow"
End Property

Public Property Get AutoSizeMovie() As Boolean
Attribute AutoSizeMovie.VB_Description = "Returns/sets a value that determines whether or not the control will change the dimensions of the image size when the window size changes."
AutoSizeMovie = PropAutoSizeMovie
End Property

Public Property Let AutoSizeMovie(ByVal Value As Boolean)
PropAutoSizeMovie = Value
If MCIWndHandle <> 0 Then
    If PropAutoSizeMovie = False Then
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOAUTOSIZEMOVIE, ByVal MCIWNDF_NOAUTOSIZEMOVIE
    Else
        SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_NOAUTOSIZEMOVIE, ByVal 0&
    End If
    Call UserControl_Resize
End If
UserControl.PropertyChanged "AutoSizeMovie"
End Property

Public Property Get TimerFreq() As Integer
Attribute TimerFreq.VB_Description = "Returns/sets the time period between position updates or between updating events."
TimerFreq = PropTimerFreq
End Property

Public Property Let TimerFreq(ByVal Value As Integer)
If Value <= 0 Then
    If MCIWndDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropTimerFreq = Value
If MCIWndHandle <> 0 Then SendMessage MCIWndHandle, MCIWNDM_SETTIMERS, PropTimerFreq, ByVal CLng(PropTimerFreq)
UserControl.PropertyChanged "TimerFreq"
End Property

Public Property Get Zoom() As Long
Attribute Zoom.VB_Description = "Returns/sets the playback movie size based on a percentage of the authored size of the file."
If MCIWndHandle <> 0 Then
    Zoom = SendMessage(MCIWndHandle, MCIWNDM_GETZOOM, 0, ByVal 0&)
Else
    Zoom = PropZoom
End If
End Property

Public Property Let Zoom(ByVal Value As Long)
If Value <= 0 Then
    If MCIWndDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropZoom = Value
If MCIWndHandle <> 0 Then
    SendMessage MCIWndHandle, MCIWNDM_SETZOOM, 0, ByVal PropZoom
    Call UserControl_Resize
End If
UserControl.PropertyChanged "Zoom"
End Property

Public Property Get Caption() As MciCaptionConstants
Attribute Caption.VB_Description = "Returns/sets the information that can be displayed in the caption bar."
Caption = PropCaption
End Property

Public Property Let Caption(ByVal Value As MciCaptionConstants)
Select Case Value
    Case MciCaptionNone, MciCaptionName, MciCaptionNamePos, MciCaptionNameMode, MciCaptionNamePosMode, MciCaptionPos, MciCaptionPosMode, MciCaptionMode
        PropCaption = Value
    Case Else
        Err.Raise 380
End Select
If MCIWndHandle <> 0 Then
    SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_SHOWALL, ByVal 0&
    Dim dwStyle As Long
    If PropCaption = MciCaptionNone Then
        Call ComCtlsChangeBorderStyle(MCIWndHandle, PropBorderStyle)
        dwStyle = GetWindowLong(MCIWndHandle, GWL_STYLE)
        If (dwStyle And WS_CAPTION) = WS_CAPTION Then SetWindowLong MCIWndHandle, GWL_STYLE, dwStyle And Not WS_CAPTION
    Else
        Call ComCtlsChangeBorderStyle(MCIWndHandle, CCBorderStyleNone)
        dwStyle = GetWindowLong(MCIWndHandle, GWL_STYLE)
        If Not (dwStyle And WS_CAPTION) = WS_CAPTION Then SetWindowLong MCIWndHandle, GWL_STYLE, dwStyle Or WS_CAPTION
    End If
    Call ComCtlsFrameChanged(MCIWndHandle)
    Select Case PropCaption
        Case MciCaptionName
            SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_SHOWNAME, ByVal MCIWNDF_SHOWNAME
        Case MciCaptionNamePos
            SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_SHOWNAME Or MCIWNDF_SHOWPOS, ByVal MCIWNDF_SHOWNAME Or MCIWNDF_SHOWPOS
        Case MciCaptionNameMode
            SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_SHOWNAME Or MCIWNDF_SHOWMODE, ByVal MCIWNDF_SHOWNAME Or MCIWNDF_SHOWMODE
        Case MciCaptionNamePosMode
            SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_SHOWNAME Or MCIWNDF_SHOWPOS Or MCIWNDF_SHOWMODE, ByVal MCIWNDF_SHOWNAME Or MCIWNDF_SHOWPOS Or MCIWNDF_SHOWMODE
        Case MciCaptionPos
            SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_SHOWPOS, ByVal MCIWNDF_SHOWPOS
        Case MciCaptionPosMode
            SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_SHOWPOS Or MCIWNDF_SHOWMODE, ByVal MCIWNDF_SHOWPOS Or MCIWNDF_SHOWMODE
        Case MciCaptionMode
            SendMessage MCIWndHandle, MCIWNDM_CHANGESTYLES, MCIWNDF_SHOWMODE, ByVal MCIWNDF_SHOWMODE
    End Select
    Call UserControl_Resize
End If
UserControl.PropertyChanged "Caption"
End Property

Private Sub CreateMCIWnd()
If MCIWndHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or MCIWNDF_NOTIFYALL
If PropCaption = MciCaptionNone Then
    Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, PropBorderStyle)
Else
    dwStyle = dwStyle Or WS_CAPTION
End If
If PropErrorDlg = False Then dwStyle = dwStyle Or MCIWNDF_NOERRORDLG
If PropRecord = True Then dwStyle = dwStyle Or MCIWNDF_RECORD
If PropPlaybar = False Then dwStyle = dwStyle Or MCIWNDF_NOPLAYBAR
If PropMenu = False Then dwStyle = dwStyle Or MCIWNDF_NOMENU
If PropAllowOpen = False Or MCIWndDesignMode = True Then dwStyle = dwStyle Or MCIWNDF_NOOPEN
If PropAutoSizeWindow = False Then dwStyle = dwStyle Or MCIWNDF_NOAUTOSIZEWINDOW
If PropAutoSizeMovie = False Then dwStyle = dwStyle Or MCIWNDF_NOAUTOSIZEMOVIE
Select Case PropCaption
    Case MciCaptionName
        dwStyle = dwStyle Or MCIWNDF_SHOWNAME
    Case MciCaptionNamePos
        dwStyle = dwStyle Or MCIWNDF_SHOWNAME Or MCIWNDF_SHOWPOS
    Case MciCaptionNameMode
        dwStyle = dwStyle Or MCIWNDF_SHOWNAME Or MCIWNDF_SHOWMODE
    Case MciCaptionNamePosMode
        dwStyle = dwStyle Or MCIWNDF_SHOWNAME Or MCIWNDF_SHOWPOS Or MCIWNDF_SHOWMODE
    Case MciCaptionPos
        dwStyle = dwStyle Or MCIWNDF_SHOWPOS
    Case MciCaptionPosMode
        dwStyle = dwStyle Or MCIWNDF_SHOWPOS Or MCIWNDF_SHOWMODE
    Case MciCaptionMode
        dwStyle = dwStyle Or MCIWNDF_SHOWMODE
End Select
MCIWndHandle = CreateWindowEx(dwExStyle, StrPtr("MCIWndClass"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
Me.Enabled = UserControl.Enabled
Me.VisualStyles = PropVisualStyles
Me.Repeat = PropRepeat
Me.TimerFreq = PropTimerFreq
Me.Zoom = PropZoom
If MCIWndDesignMode = False Then
    If MCIWndHandle <> 0 Then
        If MCIWndBackColorBrush = 0 Then MCIWndBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
        Call ComCtlsSetSubclass(MCIWndHandle, Me, 1)
        Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
    End If
Else
    If MCIWndHandle <> 0 Then
        If MCIWndBackColorBrush = 0 Then MCIWndBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
        Call ComCtlsSetSubclass(MCIWndHandle, Me, 3)
    End If
End If
End Sub

Private Sub DestroyMCIWnd()
If MCIWndHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(MCIWndHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
SendMessage MCIWndHandle, WM_CLOSE, 0, ByVal 0& ' MCIWndDestroy
MCIWndHandle = 0
If MCIWndBackColorBrush <> 0 Then
    DeleteObject MCIWndBackColorBrush
    MCIWndBackColorBrush = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
If MCIWndHandle <> 0 Then SendMessage MCIWndHandle, MCIWNDM_VALIDATEMEDIA, 0, ByVal 0&
End Sub

Public Function ShowOpen() As Boolean
Attribute ShowOpen.VB_Description = "Displays the open dialog box."
If MCIWndHandle <> 0 Then ShowOpen = CBool(SendMessage(MCIWndHandle, MCIWNDM_OPEN, 0, ByVal -1&) = 0)
End Function

Public Function ShowSave() As Boolean
Attribute ShowSave.VB_Description = "Displays the save dialog box."
If MCIWndHandle <> 0 Then ShowSave = CBool(SendMessage(MCIWndHandle, MCI_SAVE, 0, ByVal -1&) = 0)
End Function

Public Function CanSave() As Boolean
Attribute CanSave.VB_Description = "Determines if an MCI device can save."
If MCIWndHandle <> 0 Then CanSave = CBool(SendMessage(MCIWndHandle, MCIWNDM_CAN_SAVE, 0, ByVal 0&) <> 0)
End Function

Public Function Eject() As Boolean
Attribute Eject.VB_Description = "Sends a command to an MCI device to eject its media."
If MCIWndHandle <> 0 Then Eject = CBool(SendMessage(MCIWndHandle, MCIWNDM_EJECT, 0, ByVal 0&) = 0)
End Function

Public Function CanEject() As Boolean
Attribute CanEject.VB_Description = "Determines if an MCI device can eject its media."
If MCIWndHandle <> 0 Then CanEject = CBool(SendMessage(MCIWndHandle, MCIWNDM_CAN_EJECT, 0, ByVal 0&) <> 0)
End Function

Public Function PlayFrom(ByVal StartPosition As Long) As Boolean
Attribute PlayFrom.VB_Description = "Plays the content of an MCI device from the specified start to the end of the content."
If MCIWndHandle <> 0 Then PlayFrom = CBool(SendMessage(MCIWndHandle, MCIWNDM_PLAYFROM, 0, ByVal StartPosition) = 0)
End Function

Public Function PlayTo(ByVal EndPosition As Long) As Boolean
Attribute PlayTo.VB_Description = "Plays the content of an MCI device from the current position to the specified ending location."
If MCIWndHandle <> 0 Then PlayTo = CBool(SendMessage(MCIWndHandle, MCIWNDM_PLAYTO, 0, ByVal EndPosition) = 0)
End Function

Public Function PlayReverse() As Boolean
Attribute PlayReverse.VB_Description = "Plays the current content in the reverse direction, beginning at the current position and ending at the beginning of the content."
If MCIWndHandle <> 0 Then PlayReverse = CBool(SendMessage(MCIWndHandle, MCIWNDM_PLAYREVERSE, 0, ByVal 0&) = 0)
End Function

Public Function CanPlay() As Boolean
Attribute CanPlay.VB_Description = "Determines if an MCI device can play."
If MCIWndHandle <> 0 Then CanPlay = CBool(SendMessage(MCIWndHandle, MCIWNDM_CAN_PLAY, 0, ByVal 0&) <> 0)
End Function

Public Function CanRecord() As Boolean
Attribute CanRecord.VB_Description = "Determines if an MCI device supports recording."
If MCIWndHandle <> 0 Then CanRecord = CBool(SendMessage(MCIWndHandle, MCIWNDM_CAN_RECORD, 0, ByVal 0&) <> 0)
End Function

Public Function CanConfig() As Boolean
Attribute CanConfig.VB_Description = "Determines if an MCI device can display a configuration dialog box."
If MCIWndHandle <> 0 Then CanConfig = CBool(SendMessage(MCIWndHandle, MCIWNDM_CAN_CONFIG, 0, ByVal 0&) <> 0)
End Function

Public Function CanWindow() As Boolean
Attribute CanWindow.VB_Description = "Determines if an MCI device supports window-oriented MCI commands."
If MCIWndHandle <> 0 Then CanWindow = CBool(SendMessage(MCIWndHandle, MCIWNDM_CAN_WINDOW, 0, ByVal 0&) <> 0)
End Function

Public Property Get Command() As String
Attribute Command.VB_Description = "Specifies an MCI command string to execute."
Attribute Command.VB_MemberFlags = "400"
Command = MCIWndCommand
End Property

Public Property Let Command(ByVal Value As String)
MCIWndCommand = Value
If MCIWndHandle <> 0 Then SendMessage MCIWndHandle, MCIWNDM_SENDSTRING, 0, ByVal StrPtr(MCIWndCommand)
End Property

Public Property Get CommandReturn() As String
Attribute CommandReturn.VB_Description = "Returns the string from previously executed MCI command."
Attribute CommandReturn.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then
    Dim Buffer As String
    Buffer = String(256, vbNullChar)
    SendMessage MCIWndHandle, MCIWNDM_RETURNSTRING, LenB(Buffer), ByVal StrPtr(Buffer)
    CommandReturn = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Public Property Let CommandReturn(ByVal Value As String)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the device or device element (file) to be opened."
Attribute FileName.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then
    Dim Buffer As String
    Buffer = String(256, vbNullChar)
    If SendMessage(MCIWndHandle, MCIWNDM_GETFILENAME, LenB(Buffer), ByVal StrPtr(Buffer)) = 0 Then
        FileName = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End If
End If
End Property

Public Property Let FileName(ByVal Value As String)
If MCIWndHandle <> 0 Then If SendMessage(MCIWndHandle, MCIWNDM_OPEN, 0, ByVal StrPtr(Value)) <> 0 Then Err.Raise 53
End Property

Public Property Get DeviceAlias() As Long
Attribute DeviceAlias.VB_Description = "Returns the device alias of the currently open device element."
Attribute DeviceAlias.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then DeviceAlias = SendMessage(MCIWndHandle, MCIWNDM_GETALIAS, 0, ByVal 0&)
End Property

Public Property Let DeviceAlias(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get DeviceID() As Long
Attribute DeviceID.VB_Description = "Returns the device ID of the currently open device element."
Attribute DeviceID.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then DeviceID = SendMessage(MCIWndHandle, MCIWNDM_GETDEVICEID, 0, ByVal 0&)
End Property

Public Property Let DeviceID(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Device() As String
Attribute Device.VB_Description = "Returns/sets the type of the currently open device."
Attribute Device.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then
    Dim Buffer As String
    Buffer = String(256, vbNullChar)
    If SendMessage(MCIWndHandle, MCIWNDM_GETDEVICE, LenB(Buffer), ByVal StrPtr(Buffer)) = 0 Then
        Device = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End If
End If
End Property

Public Property Let Device(ByVal Value As String)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get NewDevice() As String
Attribute NewDevice.VB_Description = "Specifies a device for recording new data without an associated file."
Attribute NewDevice.VB_MemberFlags = "400"
Err.Raise Number:=394, Description:="Property is write-only"
End Property

Public Property Let NewDevice(ByVal Value As String)
If MCIWndHandle <> 0 Then SendMessage MCIWndHandle, MCIWNDM_NEW, 0, ByVal StrPtr(Value)
End Property

Public Property Get Error() As Long
Attribute Error.VB_Description = "Returns the error code generated by the last command."
Attribute Error.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then Error = SendMessage(MCIWndHandle, MCIWNDM_GETERROR, 0, ByVal 0&)
End Property

Public Property Let Error(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get ErrorString() As String
Attribute ErrorString.VB_Description = "Returns the error string returned from the last MCI command."
Attribute ErrorString.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then
    Dim Buffer As String
    Buffer = String(256, vbNullChar)
    SendMessage MCIWndHandle, MCIWNDM_GETERROR, LenB(Buffer), ByVal StrPtr(Buffer)
    ErrorString = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Public Property Let ErrorString(ByVal Value As String)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get TimeFormat() As MciFormatConstants
Attribute TimeFormat.VB_Description = "Returns/sets the current time format of an open MCI device."
Attribute TimeFormat.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then TimeFormat = SendMessage(MCIWndHandle, MCIWNDM_GETTIMEFORMAT, 0, ByVal 0&)
End Property

Public Property Let TimeFormat(ByVal Value As MciFormatConstants)
If MCIWndHandle <> 0 Then
    Dim pszFormat As String
    Select Case Value
        Case MciFormatMilliseconds
            pszFormat = "Milliseconds"
        Case MciFormatHms
            pszFormat = "Hms"
        Case MciFormatMsf
            pszFormat = "Msf"
        Case MciFormatFrames
            pszFormat = "Frames"
        Case MciFormatSmpte24
            pszFormat = "Smpte24"
        Case MciFormatSmpte25
            pszFormat = "Smpte25"
        Case MciFormatSmpte30
            pszFormat = "Smpte30"
        Case MciFormatSmpte30Drop
            pszFormat = "Smpte30Drop"
        Case MciFormatBytes
            pszFormat = "Bytes"
        Case MciFormatSamples
            pszFormat = "Samples"
        Case MciFormatTmsf
            pszFormat = "Tmsf"
        Case Else
            Err.Raise 380
    End Select
    If Not pszFormat = vbNullString Then SendMessage MCIWndHandle, MCIWNDM_SETTIMEFORMAT, 0, ByVal StrPtr(pszFormat)
End If
End Property

Public Property Get Mode() As MciModeConstants
Attribute Mode.VB_Description = "Returns the mode of the current device."
Attribute Mode.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then Mode = SendMessage(MCIWndHandle, MCIWNDM_GETMODE, 0, ByVal 0&)
End Property

Public Property Let Mode(ByVal Value As MciModeConstants)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get ModeString() As String
Attribute ModeString.VB_Description = "Returns the mode of the current device."
Attribute ModeString.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then
    Dim Buffer As String
    Buffer = String(256, vbNullChar)
    SendMessage MCIWndHandle, MCIWNDM_GETMODE, LenB(Buffer), ByVal StrPtr(Buffer)
    ModeString = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Public Property Let ModeString(ByVal Value As String)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns the current position of an open MCI device in the current time format."
Attribute Position.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then Position = SendMessage(MCIWndHandle, MCIWNDM_GETPOSITION, 0, ByVal 0&)
End Property

Public Property Let Position(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get PositionString() As String
Attribute PositionString.VB_Description = "Returns the current position of an open MCI device. If the device supports tracks, the position is returned in tracks, minutes, seconds, and frames; otherwise, it is returned as a string."
Attribute PositionString.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then
    Dim Buffer As String
    Buffer = String(256, vbNullChar)
    SendMessage MCIWndHandle, MCIWNDM_GETPOSITION, LenB(Buffer), ByVal StrPtr(Buffer)
    PositionString = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Public Property Let PositionString(ByVal Value As String)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get StartPositon() As Long
Attribute StartPositon.VB_Description = "Returns the start positon of the device element currently opened in the current time format."
Attribute StartPositon.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then StartPositon = SendMessage(MCIWndHandle, MCIWNDM_GETSTART, 0, ByVal 0&)
End Property

Public Property Let StartPositon(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Length() As Long
Attribute Length.VB_Description = "Returns the length of the device element currently opened in the current time format."
Attribute Length.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then Length = SendMessage(MCIWndHandle, MCIWNDM_GETLENGTH, 0, ByVal 0&)
End Property

Public Property Let Length(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get EndPosition() As Long
Attribute EndPosition.VB_Description = "Returns the end positon of the device element currently opened in the current time format."
Attribute EndPosition.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then EndPosition = SendMessage(MCIWndHandle, MCIWNDM_GETEND, 0, ByVal 0&)
End Property

Public Property Let EndPosition(ByVal Value As Long)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Volume() As Long
Attribute Volume.VB_Description = "Returns/sets the audio volume level of the MCI device. Specify 1000 for the normal volume level. Specify larger values for a louder volume level and smaller values for a more quiet volume level."
Attribute Volume.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then Volume = SendMessage(MCIWndHandle, MCIWNDM_GETVOLUME, 0, ByVal 0&)
End Property

Public Property Let Volume(ByVal Value As Long)
If MCIWndHandle <> 0 Then SendMessage MCIWndHandle, MCIWNDM_SETVOLUME, 0, ByVal Value
End Property

Public Property Get Speed() As Long
Attribute Speed.VB_Description = "Returns/sets the speed of the MCI device. Specify 1000 for the normal speed. Specify larger values for faster speeds and smaller values for slower speeds."
Attribute Speed.VB_MemberFlags = "400"
If MCIWndHandle <> 0 Then Speed = SendMessage(MCIWndHandle, MCIWNDM_GETSPEED, 0, ByVal 0&)
End Property

Public Property Let Speed(ByVal Value As Long)
If MCIWndHandle <> 0 Then SendMessage MCIWndHandle, MCIWNDM_SETSPEED, 0, ByVal Value
End Property

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcControlDesignMode(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_ERASEBKGND
        If MCIWndBackColorBrush <> 0 Then
            Dim RC As RECT
            GetClientRect hWnd, RC
            FillRect wParam, RC, MCIWndBackColorBrush
            WindowProcControl = 1
            Exit Function
        End If
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            MCIWndCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If MCIWndCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(MCIWndCharCodeCache And &HFFFF&)
            MCIWndCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        wParam = CIntToUInt(KeyChar)
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcControl = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(wParam)
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcControl = 0
        End If
        Exit Function
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_LBUTTONDOWN
        If GetFocus() <> hWnd Then SetFocusAPI UserControl.hWnd ' UCNoSetFocusFwd not applicable
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
                WindowProcControl = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_SYSCOMMAND
        If (wParam And &HFFF0&) = SC_MOVE Then Exit Function
    Case MM_MCINOTIFY
        RaiseEvent Notify(wParam)
    Case MM_MCISIGNAL
        RaiseEvent Signal(wParam, lParam)
    Case MCIWNDM_OPEN, MCI_SAVE
        If lParam = -1 Then ' Alias to request for a file name
            ' The intrinsic dialogs are obselete and not supported prior Windows 7.
            ' It works on Windows 7 (or above), but the dialogs are very Windows 3.1 ish.
            ' Thus it is necessary to override those messages.
            With New CommonDialog
            Dim RetVal As Long, Buffer As String, FileTypes As Variant, i As Long
            Buffer = String(256, vbNullChar)
            RetVal = GetProfileString(StrPtr("MCI EXTENSIONS"), 0, StrPtr("*.*"), StrPtr(Buffer), Len(Buffer))
            If RetVal > 0 Then
                FileTypes = Split(Left$(Buffer, RetVal), vbNullChar)
                .Filter = "MCI Files|"
                For i = 0 To UBound(FileTypes) - 1
                    .Filter = .Filter & "*." & FileTypes(i) & ";"
                Next i
                .Filter = .Filter & "|All Files|*.*"
            Else
                .Filter = "All Files|*.*"
            End If
            Select Case wMsg
                Case MCIWNDM_OPEN
                    .Flags = CdlOFNExplorer Or CdlOFNPathMustExist Or CdlOFNFileMustExist Or CdlOFNHideReadOnly Or CdlOFNNoReadOnlyReturn
                    .DialogTitle = "Open MCI File"
                    If .ShowOpen = True Then
                        WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, StrPtr(.FileName))
                    Else
                        WindowProcControl = -1
                    End If
                Case MCI_SAVE
                    .Flags = CdlOFNExplorer Or CdlOFNPathMustExist Or CdlOFNHideReadOnly Or CdlOFNOverwritePrompt
                    .DialogTitle = "Save MCI File"
                    If .ShowSave = True Then
                        WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, StrPtr(.FileName))
                    Else
                        WindowProcControl = -1
                    End If
            End Select
            End With
            Exit Function
        End If
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        RaiseEvent DblClick
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                MCIWndIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                MCIWndIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                MCIWndIsClick = True
            Case WM_MOUSEMOVE
                If MCIWndMouseOver = False And PropMouseTrack = True Then
                    MCIWndMouseOver = True
                    RaiseEvent MouseEnter
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                Select Case wMsg
                    Case WM_LBUTTONUP
                        RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_MBUTTONUP
                        RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                    Case WM_RBUTTONUP
                        RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                End Select
                If MCIWndIsClick = True Then
                    MCIWndIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If MCIWndMouseOver = True Then
            MCIWndMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case MCIWNDM_NOTIFYMODE
        If wParam = MCIWndHandle Then RaiseEvent ModeChange(lParam)
    Case MCIWNDM_NOTIFYPOS
        If wParam = MCIWndHandle Then RaiseEvent PositionChange(lParam)
    Case MCIWNDM_NOTIFYSIZE
        If wParam = MCIWndHandle Then
            Dim WndRect As RECT
            GetWindowRect MCIWndHandle, WndRect
            With UserControl
            .Extender.Move .Extender.Left, .Extender.Top, .ScaleX((WndRect.Right - WndRect.Left), vbPixels, vbContainerSize), .ScaleY((WndRect.Bottom - WndRect.Top), vbPixels, vbContainerSize)
            End With
        End If
    Case MCIWNDM_NOTIFYMEDIA
        If wParam = MCIWndHandle Then
            Dim NewFileName As String
            If lParam <> 0 Then
                Dim Buffer As String, Length As Long
                If (SendMessage(MCIWndHandle, MCIWNDM_GETSTYLES, 0, ByVal 0&) And MCIWNDF_NOTIFYANSI) = 0 Then
                    Length = lstrlen(lParam)
                    If Length > 0 Then
                        Buffer = String(Length, vbNullChar)
                        CopyMemory ByVal StrPtr(Buffer), ByVal lParam, Length * 2
                    End If
                    NewFileName = Buffer
                Else
                    Length = lstrlenA(lParam)
                    Buffer = String(Length, vbNullChar)
                    CopyMemory ByVal StrPtr(Buffer), ByVal lParam, Length
                    NewFileName = StrConv(Buffer, vbUnicode)
                End If
            End If
            RaiseEvent MediaChange(NewFileName)
        End If
    Case MCIWNDM_NOTIFYERROR
        If wParam = MCIWndHandle Then RaiseEvent Error(lParam)
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI MCIWndHandle
End Function

Private Function WindowProcControlDesignMode(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_ERASEBKGND
        WindowProcControlDesignMode = WindowProcControl(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
WindowProcControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
