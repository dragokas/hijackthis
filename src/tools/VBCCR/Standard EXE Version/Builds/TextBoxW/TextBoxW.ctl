VERSION 5.00
Begin VB.UserControl TextBoxW 
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   PropertyPages   =   "TextBoxW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "TextBoxW.ctx":0046
End
Attribute VB_Name = "TextBoxW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private TxtCharacterCasingNormal, TxtCharacterCasingUpper, TxtCharacterCasingLower
Private TxtIconNone, TxtIconInfo, TxtIconWarning, TxtIconError
Private TxtNetAddressFormatNone, TxtNetAddressFormatDNSName, TxtNetAddressFormatIPv4, TxtNetAddressFormatIPv6
Private TxtNetAddressTypeNone, TxtNetAddressTypeIPv4Address, TxtNetAddressTypeIPv4Service, TxtNetAddressTypeIPv4Network, TxtNetAddressTypeIPv6Address, TxtNetAddressTypeIPv6AddressNoScope, TxtNetAddressTypeIPv6Service, TxtNetAddressTypeIPv6ServiceNoScope, TxtNetAddressTypeIPv6Network, TxtNetAddressTypeDNSName, TxtNetAddressTypeDNSService, TxtNetAddressTypeIPAddress, TxtNetAddressTypeIPAddressNoScope, TxtNetAddressTypeIPService, TxtNetAddressTypeIPServiceNoScope, TxtNetAddressTypeIPNetwork, TxtNetAddressTypeAnyAddress, TxtNetAddressTypeAnyAddressNoScope, TxtNetAddressTypeAnyService, TxtNetAddressTypeAnyServiceNoScope
#End If
Public Enum TxtCharacterCasingConstants
TxtCharacterCasingNormal = 0
TxtCharacterCasingUpper = 1
TxtCharacterCasingLower = 2
End Enum
Private Const TTI_NONE As Long = 0
Private Const TTI_INFO As Long = 1
Private Const TTI_WARNING As Long = 2
Private Const TTI_ERROR As Long = 3
Public Enum TxtIconConstants
TxtIconNone = TTI_NONE
TxtIconInfo = TTI_INFO
TxtIconWarning = TTI_WARNING
TxtIconError = TTI_ERROR
End Enum
Private Const NET_ADDRESS_FORMAT_UNSPECIFIED As Long = 0
Private Const NET_ADDRESS_DNS_NAME As Long = 1
Private Const NET_ADDRESS_IPV4 As Long = 2
Private Const NET_ADDRESS_IPV6 As Long = 3
Public Enum TxtNetAddressFormatConstants
TxtNetAddressFormatNone = NET_ADDRESS_FORMAT_UNSPECIFIED
TxtNetAddressFormatDNSName = NET_ADDRESS_DNS_NAME
TxtNetAddressFormatIPv4 = NET_ADDRESS_IPV4
TxtNetAddressFormatIPv6 = NET_ADDRESS_IPV6
End Enum
Public Enum TxtNetAddressTypeConstants
TxtNetAddressTypeNone = 0
TxtNetAddressTypeIPv4Address = 1
TxtNetAddressTypeIPv4Service = 2
TxtNetAddressTypeIPv4Network = 3
TxtNetAddressTypeIPv6Address = 4
TxtNetAddressTypeIPv6AddressNoScope = 5
TxtNetAddressTypeIPv6Service = 6
TxtNetAddressTypeIPv6ServiceNoScope = 7
TxtNetAddressTypeIPv6Network = 8
TxtNetAddressTypeDNSName = 9
TxtNetAddressTypeDNSService = 10
TxtNetAddressTypeIPAddress = 11
TxtNetAddressTypeIPAddressNoScope = 12
TxtNetAddressTypeIPService = 13
TxtNetAddressTypeIPServiceNoScope = 14
TxtNetAddressTypeIPNetwork = 15
TxtNetAddressTypeAnyAddress = 16
TxtNetAddressTypeAnyAddressNoScope = 17
TxtNetAddressTypeAnyService = 18
TxtNetAddressTypeAnyServiceNoScope = 19
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type EDITBALLOONTIP
cbStruct As Long
pszTitle As Long
pszText As Long
iIcon As Long
End Type
Private Type NET_ADDRESS_INFO_UNSPECIFIED
Format As Integer
Data(0 To (1024 - 1)) As Byte
End Type
Private Const DNS_MAX_NAME_BUFFER_LENGTH As Long = 256
Private Type NET_ADDRESS_INFO_DNS_NAME
Format As Integer
Address(0 To ((DNS_MAX_NAME_BUFFER_LENGTH * 2) - 1)) As Byte
Port(0 To ((6 * 2) - 1)) As Byte
End Type
Private Type NET_ADDRESS_INFO_IPV4
Format As Integer
sin_family As Integer
sin_port As Integer
sin_addr As Long
sin_zero(0 To (8 - 1)) As Byte
End Type
Private Type NET_ADDRESS_INFO_IPV6
Format As Integer
sin6_family As Integer
sin6_port As Integer
sin6_flowinfoLo As Integer
sin6_flowinfoHi As Integer
sin6_addr(0 To (8 - 1)) As Integer
sin6_scope_idLo As Integer
sin6_scope_idHi As Integer
End Type
Private Type NC_ADDRESS
pAddrInfo As Long ' VarPtr(NET_ADDRESS_INFO_*)
PortNumber As Integer
PrefixLength As Byte
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event MaxText()
Attribute MaxText.VB_Description = "Occurs when the current text insertion has exceeded the maximum number of characters that can be entered in a control."
Public Event Scroll()
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Public Event ContextMenu(ByRef Handled As Boolean, ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
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
Private Declare Function InitNetworkAddressControl Lib "shell32" () As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DestroyCaret Lib "user32" () As Long
Private Declare Function DragDetect Lib "user32" (ByVal hWnd As Long, ByVal PX As Integer, ByVal PY As Integer) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const CF_UNICODETEXT As Long = 13
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_RTLREADING As Long = &H2000, WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const SB_LINELEFT As Long = 0, SB_LINERIGHT As Long = 1
Private Const SB_LINEUP As Long = 0, SB_LINEDOWN As Long = 1
Private Const SB_THUMBPOSITION As Long = 4, SB_THUMBTRACK As Long = 5
Private Const SB_HORZ As Long = 0, SB_VERT As Long = 1
Private Const SW_HIDE As Long = &H0
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_COMMAND As Long = &H111
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_UNICHAR As Long = &H109, UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_INPUTLANGCHANGE As Long = &H51
Private Const WM_IME_SETCONTEXT As Long = &H281
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
Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_COPY As Long = &H301
Private Const WM_CUT As Long = &H300
Private Const WM_PASTE As Long = &H302
Private Const WM_CLEAR As Long = &H303
Private Const WM_USER As Long = &H400
Private Const NCM_GETADDRESS As Long = (WM_USER + 1)
Private Const NCM_SETALLOWTYPE As Long = (WM_USER + 2)
Private Const NCM_GETALLOWTYPE As Long = (WM_USER + 3)
Private Const NCM_DISPLAYERRORTIP As Long = (WM_USER + 4)
Private Const NET_STRING_IPV4_ADDRESS As Long = &H1
Private Const NET_STRING_IPV4_SERVICE As Long = &H2
Private Const NET_STRING_IPV4_NETWORK As Long = &H4
Private Const NET_STRING_IPV6_ADDRESS As Long = &H8
Private Const NET_STRING_IPV6_ADDRESS_NO_SCOPE As Long = &H10
Private Const NET_STRING_IPV6_SERVICE As Long = &H20
Private Const NET_STRING_IPV6_SERVICE_NO_SCOPE As Long = &H40
Private Const NET_STRING_IPV6_NETWORK As Long = &H80
Private Const NET_STRING_NAMED_ADDRESS As Long = &H100
Private Const NET_STRING_NAMED_SERVICE As Long = &H200
Private Const NET_STRING_IP_ADDRESS As Long = (NET_STRING_IPV4_ADDRESS Or NET_STRING_IPV6_ADDRESS)
Private Const NET_STRING_IP_ADDRESS_NO_SCOPE As Long = (NET_STRING_IPV4_ADDRESS Or NET_STRING_IPV6_ADDRESS_NO_SCOPE)
Private Const NET_STRING_IP_SERVICE As Long = (NET_STRING_IPV4_SERVICE Or NET_STRING_IPV6_SERVICE)
Private Const NET_STRING_IP_SERVICE_NO_SCOPE As Long = (NET_STRING_IPV4_SERVICE Or NET_STRING_IPV6_SERVICE_NO_SCOPE)
Private Const NET_STRING_IP_NETWORK As Long = (NET_STRING_IPV4_NETWORK Or NET_STRING_IPV6_NETWORK)
Private Const NET_STRING_ANY_ADDRESS As Long = (NET_STRING_NAMED_ADDRESS Or NET_STRING_IP_ADDRESS)
Private Const NET_STRING_ANY_ADDRESS_NO_SCOPE As Long = (NET_STRING_NAMED_ADDRESS Or NET_STRING_IP_ADDRESS_NO_SCOPE)
Private Const NET_STRING_ANY_SERVICE As Long = (NET_STRING_NAMED_SERVICE Or NET_STRING_IP_SERVICE)
Private Const NET_STRING_ANY_SERVICE_NO_SCOPE As Long = (NET_STRING_NAMED_SERVICE Or NET_STRING_IP_SERVICE_NO_SCOPE)
Private Const EM_SETREADONLY As Long = &HCF, ES_READONLY As Long = &H800
Private Const EM_GETSEL As Long = &HB0
Private Const EM_SETSEL As Long = &HB1
Private Const EM_LINESCROLL As Long = &HB6
Private Const EM_SCROLLCARET As Long = &HB7
Private Const EM_REPLACESEL As Long = &HC2
Private Const EM_GETPASSWORDCHAR As Long = &HD2
Private Const EM_SETPASSWORDCHAR As Long = &HCC
Private Const EM_GETLIMITTEXT As Long = &HD5
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETLIMITTEXT As Long = EM_LIMITTEXT
Private Const EM_GETMODIFY As Long = &HB8
Private Const EM_SETMODIFY As Long = &HB9
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_GETLINE As Long = &HC4
Private Const EM_UNDO As Long = &HC7
Private Const EM_CANUNDO As Long = &HC6
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_EMPTYUNDOBUFFER As Long = &HCD
Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_GETMARGINS As Long = &HD4
Private Const EM_SETMARGINS As Long = &HD3
Private Const EM_POSFROMCHAR As Long = &HD6
Private Const EM_CHARFROMPOS As Long = &HD7
Private Const ECM_FIRST As Long = &H1500
Private Const EM_SETCUEBANNER As Long = (ECM_FIRST + 1)
Private Const EM_GETCUEBANNER As Long = (ECM_FIRST + 2)
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)
Private Const EN_CHANGE As Long = &H300
Private Const EN_UPDATE As Long = &H400
Private Const EN_MAXTEXT As Long = &H501
Private Const EN_HSCROLL As Long = &H601
Private Const EN_VSCROLL As Long = &H602
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const ES_AUTOVSCROLL As Long = &H40
Private Const ES_NUMBER As Long = &H2000
Private Const ES_NOHIDESEL As Long = &H100
Private Const ES_LEFT As Long = &H0
Private Const ES_CENTER As Long = &H1
Private Const ES_RIGHT As Long = &H2
Private Const ES_MULTILINE As Long = &H4
Private Const ES_UPPERCASE As Long = &H8
Private Const ES_LOWERCASE As Long = &H10
Private Const ES_PASSWORD As Long = &H20
Private Const EC_LEFTMARGIN As Long = &H1
Private Const EC_RIGHTMARGIN As Long = &H2
Private Const EC_USEFONTINFO As Long = &HFFFF&
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IOleControlVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private TextBoxHandle As Long
Private TextBoxFontHandle As Long
Private TextBoxIMCHandle As Long
Private TextBoxCharCodeCache As Long
Private TextBoxAutoDragInSel As Boolean, TextBoxAutoDragIsActive As Boolean
Private TextBoxIsClick As Boolean
Private TextBoxMouseOver As Boolean
Private TextBoxDesignMode As Boolean
Private TextBoxChangeFrozen As Boolean
Private TextBoxNetAddressFormat As TxtNetAddressFormatConstants
Private TextBoxNetAddressString As String
Private TextBoxNetAddressPortNumber As Integer
Private TextBoxNetAddressPrefixLength As Byte
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropOLEDragDropScroll As Boolean
Private PropOLEDropMode As VBRUN.OLEDropConstants
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBorderStyle As CCBorderStyleConstants
Private PropText As String
Private PropAlignment As VBRUN.AlignmentConstants
Private PropAllowOnlyNumbers As Boolean
Private PropLocked As Boolean
Private PropHideSelection As Boolean
Private PropPasswordChar As Integer
Private PropUseSystemPasswordChar As Boolean
Private PropMultiLine As Boolean
Private PropMaxLength As Long
Private PropScrollBars As VBRUN.ScrollBarConstants
Private PropCueBanner As String
Private PropCharacterCasing As TxtCharacterCasingConstants
Private PropWantReturn As Boolean
Private PropIMEMode As CCIMEModeConstants
Private PropNetAddressValidator As Boolean
Private PropNetAddressType As TxtNetAddressTypeConstants
Private PropAllowOverType As Boolean
Private PropOverTypeMode As Boolean

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

Private Sub IOleControlVB_GetControlInfo(ByRef Handled As Boolean, ByRef AccelCount As Integer, ByRef AccelTable As Long, ByRef Flags As Long)
If PropWantReturn = True And PropMultiLine = True Then
    Flags = CTRLINFO_EATS_RETURN
    Handled = True
End If
End Sub

Private Sub IOleControlVB_OnMnemonic(ByRef Handled As Boolean, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
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
Call ComCtlsInitCC(ICC_STANDARD_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfaceControl)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
TextBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropOLEDragMode = vbOLEDragManual
PropOLEDragDropScroll = True
PropOLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = CCBorderStyleSunken
PropText = Ambient.DisplayName
If PropRightToLeft = False Then PropAlignment = vbLeftJustify Else PropAlignment = vbRightJustify
PropAllowOnlyNumbers = False
PropLocked = False
PropHideSelection = True
PropPasswordChar = 0
PropUseSystemPasswordChar = False
PropMultiLine = False
PropMaxLength = 0
PropScrollBars = vbSBNone
PropCueBanner = vbNullString
PropCharacterCasing = TxtCharacterCasingNormal
PropWantReturn = False
PropIMEMode = CCIMEModeNoControl
PropNetAddressValidator = False
PropNetAddressType = TxtNetAddressTypeNone
PropAllowOverType = False
PropOverTypeMode = False
Call CreateTextBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
TextBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbWindowBackground)
Me.ForeColor = .ReadProperty("ForeColor", vbWindowText)
Me.Enabled = .ReadProperty("Enabled", True)
PropOLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
PropOLEDragDropScroll = .ReadProperty("OLEDragDropScroll", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSunken)
PropText = VarToStr(.ReadProperty("Text", vbNullString))
PropAlignment = .ReadProperty("Alignment", vbLeftJustify)
PropAllowOnlyNumbers = .ReadProperty("AllowOnlyNumbers", False)
PropLocked = .ReadProperty("Locked", False)
PropHideSelection = .ReadProperty("HideSelection", True)
Dim VarValue As Variant
VarValue = .ReadProperty("PasswordChar", 0)
If VarType(VarValue) = vbString Then ' Compatibility
    If Len(VarValue) > 0 Then PropPasswordChar = AscW(VarValue) Else PropPasswordChar = 0
Else
    PropPasswordChar = VarValue
End If
PropUseSystemPasswordChar = .ReadProperty("UseSystemPasswordChar", False)
PropMultiLine = .ReadProperty("MultiLine", False)
PropMaxLength = .ReadProperty("MaxLength", 0)
PropScrollBars = .ReadProperty("ScrollBars", vbSBNone)
PropCueBanner = VarToStr(.ReadProperty("CueBanner", vbNullString))
PropCharacterCasing = .ReadProperty("CharacterCasing", TxtCharacterCasingNormal)
PropWantReturn = .ReadProperty("WantReturn", False)
PropIMEMode = .ReadProperty("IMEMode", CCIMEModeNoControl)
PropNetAddressValidator = .ReadProperty("NetAddressValidator", False)
PropNetAddressType = .ReadProperty("NetAddressType", TxtNetAddressTypeNone)
PropAllowOverType = .ReadProperty("AllowOverType", False)
PropOverTypeMode = .ReadProperty("OverTypeMode", False)
End With
Call CreateTextBox
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbWindowBackground
.WriteProperty "ForeColor", Me.ForeColor, vbWindowText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDragDropScroll", PropOLEDragDropScroll, True
.WriteProperty "OLEDropMode", PropOLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSunken
.WriteProperty "Text", StrToVar(PropText), vbNullString
.WriteProperty "Alignment", PropAlignment, vbLeftJustify
.WriteProperty "AllowOnlyNumbers", PropAllowOnlyNumbers, False
.WriteProperty "Locked", PropLocked, False
.WriteProperty "HideSelection", PropHideSelection, True
.WriteProperty "PasswordChar", PropPasswordChar, 0
.WriteProperty "UseSystemPasswordChar", PropUseSystemPasswordChar, False
.WriteProperty "MultiLine", PropMultiLine, False
.WriteProperty "MaxLength", PropMaxLength, 0
.WriteProperty "ScrollBars", PropScrollBars, vbSBNone
.WriteProperty "CueBanner", StrToVar(PropCueBanner), vbNullString
.WriteProperty "CharacterCasing", PropCharacterCasing, TxtCharacterCasingNormal
.WriteProperty "WantReturn", PropWantReturn, False
.WriteProperty "IMEMode", PropIMEMode, CCIMEModeNoControl
.WriteProperty "NetAddressValidator", PropNetAddressValidator, False
.WriteProperty "NetAddressType", PropNetAddressType, TxtNetAddressTypeNone
.WriteProperty "AllowOverType", PropAllowOverType, False
.WriteProperty "OverTypeMode", PropOverTypeMode, False
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
If PropOLEDragMode = vbOLEDragAutomatic And TextBoxAutoDragIsActive = True And Effect = vbDropEffectMove Then
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_CLEAR, 0, ByVal 0&
End If
RaiseEvent OLECompleteDrag(Effect)
TextBoxAutoDragIsActive = False
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Text As String
If PropOLEDropMode = vbOLEDropAutomatic Then
    If Data.GetFormat(CF_UNICODETEXT) = True Then
        Text = Data.GetData(CF_UNICODETEXT) & vbNullChar
        Text = Left$(Text, InStr(Text, vbNullChar) - 1)
        Effect = vbDropEffectMove
    ElseIf Data.GetFormat(vbCFText) = True Then
        Text = Data.GetData(vbCFText)
        Effect = vbDropEffectMove
    Else
        Effect = vbDropEffectNone
    End If
End If
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
If PropOLEDropMode = vbOLEDropAutomatic Then
    If Not Effect = vbDropEffectNone And Not Text = vbNullString Then
        Me.Refresh
        If TextBoxHandle <> 0 Then
            Dim CharPos As Long
            CharPos = CIntToUInt(LoWord(SendMessage(TextBoxHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(X, Y))))
            If TextBoxAutoDragIsActive = True Then
                TextBoxAutoDragIsActive = False
                Dim SelStart As Long, SelEnd As Long
                SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                If CharPos >= SelStart And CharPos <= SelEnd Then
                    Effect = vbDropEffectNone
                    Exit Sub
                End If
                If SelStart < CharPos Then CharPos = CharPos - (SelEnd - SelStart)
                If Effect = vbDropEffectMove Then SendMessage TextBoxHandle, WM_CLEAR, 0, ByVal 0&
            Else
                If GetFocus() <> TextBoxHandle Then SetFocusAPI UserControl.hWnd
            End If
            SendMessage TextBoxHandle, EM_SETSEL, CharPos, ByVal CharPos
            SendMessage TextBoxHandle, EM_REPLACESEL, 1, ByVal StrPtr(Text)
            SendMessage TextBoxHandle, EM_SETSEL, CharPos, ByVal (CharPos + Len(Text))
        End If
    End If
End If
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If PropOLEDropMode = vbOLEDropAutomatic Then
    If Data.GetFormat(CF_UNICODETEXT) = True Or Data.GetFormat(vbCFText) = True Then Effect = vbDropEffectMove Else Effect = vbDropEffectNone
End If
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
If TextBoxHandle <> 0 Then
    If State = vbOver And Not Effect = vbDropEffectNone Then
        If PropOLEDragDropScroll = True Then
            Dim RC As RECT
            GetWindowRect TextBoxHandle, RC
            Dim dwStyle As Long
            dwStyle = GetWindowLong(TextBoxHandle, GWL_STYLE)
            If (dwStyle And WS_HSCROLL) = WS_HSCROLL Then
                If Abs(X) < (16 * PixelsPerDIP_X()) Then
                    SendMessage TextBoxHandle, WM_HSCROLL, SB_LINELEFT, ByVal 0&
                ElseIf Abs(X - (RC.Right - RC.Left)) < (16 * PixelsPerDIP_X()) Then
                    SendMessage TextBoxHandle, WM_HSCROLL, SB_LINERIGHT, ByVal 0&
                End If
            End If
            If (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                If Abs(Y) < (16 * PixelsPerDIP_Y()) Then
                    SendMessage TextBoxHandle, WM_VSCROLL, SB_LINEUP, ByVal 0&
                ElseIf Abs(Y - (RC.Bottom - RC.Top)) < (16 * PixelsPerDIP_Y()) Then
                    SendMessage TextBoxHandle, WM_VSCROLL, SB_LINEDOWN, ByVal 0&
                End If
            End If
        End If
    End If
    If PropOLEDropMode = vbOLEDropAutomatic Then
        If State = vbOver And Not Effect = vbDropEffectNone Then
            Dim CharPos As Long, CaretPos As Long
            CharPos = CIntToUInt(LoWord(SendMessage(TextBoxHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(X, Y))))
            CaretPos = SendMessage(TextBoxHandle, EM_POSFROMCHAR, CharPos, ByVal 0&)
            If CaretPos > -1 Then
                Dim hDC As Long, Size As SIZEAPI
                hDC = GetDC(TextBoxHandle)
                SelectObject hDC, TextBoxFontHandle
                GetTextExtentPoint32 hDC, StrPtr("|"), 1, Size
                ReleaseDC TextBoxHandle, hDC
                CreateCaret TextBoxHandle, 0, 0, Size.CY
                SetCaretPos LoWord(CaretPos), HiWord(CaretPos)
                ShowCaret TextBoxHandle
            Else
                If GetFocus() <> TextBoxHandle Then
                    DestroyCaret
                Else
                    Me.Refresh
                End If
            End If
        ElseIf State = vbLeave Then
            If GetFocus() <> TextBoxHandle Then
                DestroyCaret
            Else
                Me.Refresh
            End If
        End If
    End If
End If
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
If PropOLEDragMode = vbOLEDragAutomatic Then
    Dim Text As String
    Text = Me.SelText
    Data.SetData StrToVar(Text & vbNullChar), CF_UNICODETEXT
    Data.SetData Text, vbCFText
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    TextBoxAutoDragIsActive = True
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then TextBoxAutoDragIsActive = False
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If TextBoxHandle <> 0 Then MoveWindow TextBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfaceControl)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyTextBox
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
hWnd = TextBoxHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
If NewFont Is Nothing Then Set NewFont = Ambient.Font
Dim OldFontHandle As Long
Set PropFont = NewFont
OldFontHandle = TextBoxFontHandle
TextBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_SETFONT, TextBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = TextBoxFontHandle
TextBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_SETFONT, TextBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If TextBoxHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles TextBoxHandle
    Else
        RemoveVisualStyles TextBoxHandle
    End If
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
UserControl.BackColor = Value
Me.Refresh
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
UserControl.ForeColor = Value
Me.Refresh
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If TextBoxHandle <> 0 Then EnableWindow TextBoxHandle, IIf(Value = True, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDragMode() As VBRUN.OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDragMode = PropOLEDragMode
End Property

Public Property Let OLEDragMode(ByVal Value As VBRUN.OLEDragConstants)
Select Case Value
    Case vbOLEDragManual, vbOLEDragAutomatic
        PropOLEDragMode = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "OLEDragMode"
End Property

Public Property Get OLEDragDropScroll() As Boolean
Attribute OLEDragDropScroll.VB_Description = "Returns/Sets whether this object will scroll during an OLE drag/drop operation."
OLEDragDropScroll = PropOLEDragDropScroll
End Property

Public Property Let OLEDragDropScroll(ByVal Value As Boolean)
PropOLEDragDropScroll = Value
UserControl.PropertyChanged "OLEDragDropScroll"
End Property

Public Property Get OLEDropMode() As VBRUN.OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
OLEDropMode = PropOLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As VBRUN.OLEDropConstants)
Select Case Value
    Case vbOLEDropNone, vbOLEDropManual, vbOLEDropAutomatic
        PropOLEDropMode = Value
        UserControl.OLEDropMode = IIf(PropOLEDropMode = vbOLEDropAutomatic, vbOLEDropManual, Value)
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
If TextBoxDesignMode = False Then Call RefreshMousePointer
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
        If TextBoxDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If TextBoxDesignMode = False Then Call RefreshMousePointer
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

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
Attribute RightToLeft.VB_UserMemId = -611
RightToLeft = PropRightToLeft
End Property

Public Property Let RightToLeft(ByVal Value As Boolean)
PropRightToLeft = Value
UserControl.RightToLeft = PropRightToLeft
Call ComCtlsCheckRightToLeft(PropRightToLeft, UserControl.RightToLeft, PropRightToLeftMode)
Dim dwMask As Long
If PropRightToLeft = True Then dwMask = WS_EX_RTLREADING Or WS_EX_LEFTSCROLLBAR
If TextBoxHandle <> 0 Then
    Call ComCtlsSetRightToLeft(TextBoxHandle, dwMask)
    If PropRightToLeft = False Then
        If PropAlignment = vbRightJustify Then Me.Alignment = vbLeftJustify
    Else
        If PropAlignment = vbLeftJustify Then Me.Alignment = vbRightJustify
    End If
End If
UserControl.PropertyChanged "RightToLeft"
End Property

Public Property Get RightToLeftMode() As CCRightToLeftModeConstants
Attribute RightToLeftMode.VB_Description = "Returns/sets the right-to-left mode."
RightToLeftMode = PropRightToLeftMode
End Property

Public Property Let RightToLeftMode(ByVal Value As CCRightToLeftModeConstants)
Select Case Value
    Case CCRightToLeftModeNoControl, CCRightToLeftModeVBAME, CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
        PropRightToLeftMode = Value
    Case Else
        Err.Raise 380
End Select
Me.RightToLeft = PropRightToLeft
UserControl.PropertyChanged "RightToLeftMode"
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
If TextBoxHandle <> 0 Then Call ComCtlsChangeBorderStyle(TextBoxHandle, PropBorderStyle)
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_ProcData.VB_Invoke_Property = "PPTextBoxWText"
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "123c"
If TextBoxHandle <> 0 Then
    Text = String$(SendMessage(TextBoxHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage TextBoxHandle, WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
Else
    Text = PropText
End If
End Property

Public Property Let Text(ByVal Value As String)
If PropMaxLength > 0 Then Value = Left$(Value, PropMaxLength)
Dim Changed As Boolean
Changed = CBool(Me.Text <> Value)
PropText = Value
If TextBoxHandle <> 0 Then
    TextBoxChangeFrozen = True
    SendMessage TextBoxHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
    TextBoxChangeFrozen = False
End If
UserControl.PropertyChanged "Text"
If Changed = True Then
    On Error Resume Next
    UserControl.Extender.DataChanged = True
    On Error GoTo 0
    RaiseEvent Change
End If
End Property

Public Property Get Default() As String
Attribute Default.VB_UserMemId = 0
Attribute Default.VB_MemberFlags = "40"
Default = Me.Text
End Property

Public Property Let Default(ByVal Value As String)
Me.Text = Value
End Property

Public Property Get Alignment() As VBRUN.AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment."
Alignment = PropAlignment
End Property

Public Property Let Alignment(ByVal Value As VBRUN.AlignmentConstants)
Select Case Value
    Case vbLeftJustify, vbCenter, vbRightJustify
        PropAlignment = Value
    Case Else
        Err.Raise 380
End Select
If TextBoxHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TextBoxHandle, GWL_STYLE)
    If (dwStyle And ES_LEFT) = ES_LEFT Then dwStyle = dwStyle And Not ES_LEFT
    If (dwStyle And ES_CENTER) = ES_CENTER Then dwStyle = dwStyle And Not ES_CENTER
    If (dwStyle And ES_RIGHT) = ES_RIGHT Then dwStyle = dwStyle And Not ES_RIGHT
    Select Case PropAlignment
        Case vbLeftJustify
            dwStyle = dwStyle Or ES_LEFT
        Case vbCenter
            dwStyle = dwStyle Or ES_CENTER
        Case vbRightJustify
            dwStyle = dwStyle Or ES_RIGHT
    End Select
    SetWindowLong TextBoxHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "Alignment"
End Property

Public Property Get AllowOnlyNumbers() As Boolean
Attribute AllowOnlyNumbers.VB_Description = "Returns/sets a value indicating if only numbers are allowed to be entered."
AllowOnlyNumbers = PropAllowOnlyNumbers
End Property

Public Property Let AllowOnlyNumbers(ByVal Value As Boolean)
PropAllowOnlyNumbers = Value
If TextBoxHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TextBoxHandle, GWL_STYLE)
    If PropAllowOnlyNumbers = True Then
        If Not (dwStyle And ES_NUMBER) = ES_NUMBER Then dwStyle = dwStyle Or ES_NUMBER
    Else
        If (dwStyle And ES_NUMBER) = ES_NUMBER Then dwStyle = dwStyle And Not ES_NUMBER
    End If
    SetWindowLong TextBoxHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "AllowOnlyNumbers"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents can be edited."
If TextBoxHandle <> 0 Then
    Locked = CBool((GetWindowLong(TextBoxHandle, GWL_STYLE) And ES_READONLY) <> 0)
Else
    Locked = PropLocked
End If
End Property

Public Property Let Locked(ByVal Value As Boolean)
PropLocked = Value
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Locked"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value indicating if the selection in an edit control is hidden when the control loses focus."
HideSelection = PropHideSelection
End Property

Public Property Let HideSelection(ByVal Value As Boolean)
PropHideSelection = Value
If TextBoxHandle <> 0 Then Call ReCreateTextBox
UserControl.PropertyChanged "HideSelection"
End Property

Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
If TextBoxHandle <> 0 Then
    PasswordChar = ChrW(SendMessage(TextBoxHandle, EM_GETPASSWORDCHAR, 0, ByVal 0&))
Else
    PasswordChar = ChrW(PropPasswordChar)
End If
End Property

Public Property Let PasswordChar(ByVal Value As String)
If PropUseSystemPasswordChar = True Then Exit Property
If Value = vbNullString Or Len(Value) = 0 Then
    PropPasswordChar = 0
ElseIf Len(Value) = 1 Then
    PropPasswordChar = AscW(Value)
Else
    If TextBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If TextBoxHandle <> 0 Then
    SendMessage TextBoxHandle, EM_SETPASSWORDCHAR, PropPasswordChar, ByVal 0&
    Me.Refresh
End If
UserControl.PropertyChanged "PasswordChar"
End Property

Public Property Get UseSystemPasswordChar() As Boolean
Attribute UseSystemPasswordChar.VB_Description = "Returns/sets a value indicating if the default system password character is used. This property has precedence over the password char property."
UseSystemPasswordChar = PropUseSystemPasswordChar
End Property

Public Property Let UseSystemPasswordChar(ByVal Value As Boolean)
PropUseSystemPasswordChar = Value
If TextBoxHandle <> 0 Then Call ReCreateTextBox
UserControl.PropertyChanged "UseSystemPasswordChar"
End Property

Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
MultiLine = PropMultiLine
End Property

Public Property Let MultiLine(ByVal Value As Boolean)
PropMultiLine = Value
If TextBoxHandle <> 0 Then Call ReCreateTextBox
UserControl.PropertyChanged "MultiLine"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
MaxLength = PropMaxLength
End Property

Public Property Let MaxLength(ByVal Value As Long)
If Value < 0 Then
    If TextBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropMaxLength = Value
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETLIMITTEXT, PropMaxLength, ByVal 0&
UserControl.PropertyChanged "MaxLength"
End Property

Public Property Get ScrollBars() As VBRUN.ScrollBarConstants
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
ScrollBars = PropScrollBars
End Property

Public Property Let ScrollBars(ByVal Value As VBRUN.ScrollBarConstants)
Select Case Value
    Case vbSBNone, vbHorizontal, vbVertical, vbBoth
        PropScrollBars = Value
        If TextBoxHandle <> 0 Then Call ReCreateTextBox
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "ScrollBars"
End Property

Public Property Get CueBanner() As String
Attribute CueBanner.VB_Description = "Returns/sets the textual cue, or tip, that is displayed to prompt the user for information. Only applicable if the multi line property is set to false. Requires comctl32.dll version 6.0 or higher."
CueBanner = PropCueBanner
End Property

Public Property Let CueBanner(ByVal Value As String)
PropCueBanner = Value
If TextBoxHandle <> 0 And PropMultiLine = False And ComCtlsSupportLevel() >= 1 Then SendMessage TextBoxHandle, EM_SETCUEBANNER, 0, ByVal StrPtr(PropCueBanner)
UserControl.PropertyChanged "CueBanner"
End Property

Public Property Get CharacterCasing() As TxtCharacterCasingConstants
Attribute CharacterCasing.VB_Description = "Returns/sets a value indicating if the text box modifies the case of characters as they are typed."
CharacterCasing = PropCharacterCasing
End Property

Public Property Let CharacterCasing(ByVal Value As TxtCharacterCasingConstants)
Select Case Value
    Case TxtCharacterCasingNormal, TxtCharacterCasingUpper, TxtCharacterCasingLower
        PropCharacterCasing = Value
    Case Else
        Err.Raise 380
End Select
If TextBoxHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TextBoxHandle, GWL_STYLE)
    If (dwStyle And ES_UPPERCASE) = ES_UPPERCASE Then dwStyle = dwStyle And Not ES_UPPERCASE
    If (dwStyle And ES_LOWERCASE) = ES_LOWERCASE Then dwStyle = dwStyle And Not ES_LOWERCASE
    Select Case PropCharacterCasing
        Case TxtCharacterCasingUpper
            dwStyle = dwStyle Or ES_UPPERCASE
        Case TxtCharacterCasingLower
            dwStyle = dwStyle Or ES_LOWERCASE
    End Select
    SetWindowLong TextBoxHandle, GWL_STYLE, dwStyle
    If TextBoxDesignMode = True Then
        SendMessage TextBoxHandle, WM_SETTEXT, 0, ByVal 0&
        SendMessage TextBoxHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
    End If
End If
UserControl.PropertyChanged "CharacterCasing"
End Property

Public Property Get WantReturn() As Boolean
Attribute WantReturn.VB_Description = "Returns/sets a value that determines when the user presses RETURN to perform the default button or to advance to the next line. This property applies only to a multiline text box and when there is any default button on the form."
WantReturn = PropWantReturn
End Property

Public Property Let WantReturn(ByVal Value As Boolean)
If PropWantReturn = Value Then Exit Property
PropWantReturn = Value
If TextBoxHandle <> 0 And TextBoxDesignMode = False Then
    ' It is not possible (in VB6) to achieve this when specifying ES_WANTRETURN.
    Call OnControlInfoChanged(Me, CBool(GetFocus() = TextBoxHandle))
End If
UserControl.PropertyChanged "WantReturn"
End Property

Public Property Get IMEMode() As CCIMEModeConstants
Attribute IMEMode.VB_Description = "Returns/sets the Input Method Editor (IME) mode."
IMEMode = PropIMEMode
End Property

Public Property Let IMEMode(ByVal Value As CCIMEModeConstants)
Select Case Value
    Case CCIMEModeNoControl, CCIMEModeOn, CCIMEModeOff, CCIMEModeDisable, CCIMEModeHiragana, CCIMEModeKatakana, CCIMEModeKatakanaHalf, CCIMEModeAlphaFull, CCIMEModeAlpha, CCIMEModeHangulFull, CCIMEModeHangul
        PropIMEMode = Value
    Case Else
        Err.Raise 380
End Select
If TextBoxHandle <> 0 And TextBoxDesignMode = False Then
    If GetFocus() = TextBoxHandle Then Call ComCtlsSetIMEMode(TextBoxHandle, TextBoxIMCHandle, PropIMEMode)
End If
UserControl.PropertyChanged "IMEMode"
End Property

Public Property Get NetAddressValidator() As Boolean
Attribute NetAddressValidator.VB_Description = "Returns/sets a value that indicates if the content of the control represents a network address, which you can use to input and validate the format of IPv4, IPv6 and named DNS addresses. Requires comctl32.dll version 6.1 or higher."
NetAddressValidator = PropNetAddressValidator
End Property

Public Property Let NetAddressValidator(ByVal Value As Boolean)
PropNetAddressValidator = Value
If TextBoxHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    TextBoxNetAddressFormat = TxtNetAddressFormatNone
    TextBoxNetAddressString = vbNullString
    TextBoxNetAddressPortNumber = 0
    TextBoxNetAddressPrefixLength = 0
    Call ReCreateTextBox
End If
UserControl.PropertyChanged "NetAddressValidator"
End Property

Public Property Get NetAddressType() As TxtNetAddressTypeConstants
Attribute NetAddressType.VB_Description = "Returns/sets a value which represents a network address type, which will be used as a validation mask. Requires comctl32.dll version 6.1 or higher."
NetAddressType = PropNetAddressType
End Property

Public Property Let NetAddressType(ByVal Value As TxtNetAddressTypeConstants)
Select Case Value
    Case TxtNetAddressTypeNone, TxtNetAddressTypeIPv4Address, TxtNetAddressTypeIPv4Service, TxtNetAddressTypeIPv4Network, TxtNetAddressTypeIPv6Address, TxtNetAddressTypeIPv6AddressNoScope, TxtNetAddressTypeIPv6Service, TxtNetAddressTypeIPv6ServiceNoScope, TxtNetAddressTypeIPv6Network, TxtNetAddressTypeDNSName, TxtNetAddressTypeDNSService, TxtNetAddressTypeIPAddress, TxtNetAddressTypeIPAddressNoScope, TxtNetAddressTypeIPService, TxtNetAddressTypeIPServiceNoScope, TxtNetAddressTypeIPNetwork, TxtNetAddressTypeAnyAddress, TxtNetAddressTypeAnyAddressNoScope, TxtNetAddressTypeAnyService, TxtNetAddressTypeAnyServiceNoScope
        PropNetAddressType = Value
    Case Else
        Err.Raise 380
End Select
If TextBoxHandle <> 0 And PropNetAddressValidator = True And ComCtlsSupportLevel() >= 2 Then
    Dim AddrMask As Long
    Select Case PropNetAddressType
        Case TxtNetAddressTypeNone
            AddrMask = 0
        Case TxtNetAddressTypeIPv4Address
            AddrMask = NET_STRING_IPV4_ADDRESS
        Case TxtNetAddressTypeIPv4Service
            AddrMask = NET_STRING_IPV4_SERVICE
        Case TxtNetAddressTypeIPv4Network
            AddrMask = NET_STRING_IPV4_NETWORK
        Case TxtNetAddressTypeIPv6Address
            AddrMask = NET_STRING_IPV6_ADDRESS
        Case TxtNetAddressTypeIPv6AddressNoScope
            AddrMask = NET_STRING_IPV6_ADDRESS_NO_SCOPE
        Case TxtNetAddressTypeIPv6Service
            AddrMask = NET_STRING_IPV6_SERVICE
        Case TxtNetAddressTypeIPv6ServiceNoScope
            AddrMask = NET_STRING_IPV6_SERVICE_NO_SCOPE
        Case TxtNetAddressTypeIPv6Network
            AddrMask = NET_STRING_IPV6_NETWORK
        Case TxtNetAddressTypeDNSName
            AddrMask = NET_STRING_NAMED_ADDRESS
        Case TxtNetAddressTypeDNSService
            AddrMask = NET_STRING_NAMED_SERVICE
        Case TxtNetAddressTypeIPAddress
            AddrMask = NET_STRING_IP_ADDRESS
        Case TxtNetAddressTypeIPAddressNoScope
            AddrMask = NET_STRING_IP_ADDRESS_NO_SCOPE
        Case TxtNetAddressTypeIPService
            AddrMask = NET_STRING_IP_SERVICE
        Case TxtNetAddressTypeIPServiceNoScope
            AddrMask = NET_STRING_IP_SERVICE_NO_SCOPE
        Case TxtNetAddressTypeIPNetwork
            AddrMask = NET_STRING_IP_NETWORK
        Case TxtNetAddressTypeAnyAddress
            AddrMask = NET_STRING_ANY_ADDRESS
        Case TxtNetAddressTypeAnyAddressNoScope
            AddrMask = NET_STRING_ANY_ADDRESS_NO_SCOPE
        Case TxtNetAddressTypeAnyService
            AddrMask = NET_STRING_ANY_SERVICE
        Case TxtNetAddressTypeAnyServiceNoScope
            AddrMask = NET_STRING_ANY_SERVICE_NO_SCOPE
    End Select
    SendMessage TextBoxHandle, NCM_SETALLOWTYPE, AddrMask, ByVal 0&
End If
UserControl.PropertyChanged "NetAddressType"
End Property

Public Property Get AllowOverType() As Boolean
Attribute AllowOverType.VB_Description = "Returns/sets a value indicating if overtype mode is allowed to be activated."
AllowOverType = PropAllowOverType
End Property

Public Property Let AllowOverType(ByVal Value As Boolean)
PropAllowOverType = Value
If PropAllowOverType = False Then Me.OverTypeMode = False
UserControl.PropertyChanged "AllowOverType"
End Property

Public Property Get OverTypeMode() As Boolean
Attribute OverTypeMode.VB_Description = "Returns/sets a value indicating if overtype mode is active. In overtype mode, the characters you type replace existing characters one by one."
OverTypeMode = PropOverTypeMode
End Property

Public Property Let OverTypeMode(ByVal Value As Boolean)
If PropOverTypeMode = Value Then Exit Property
If PropAllowOverType = True Then PropOverTypeMode = Value Else PropOverTypeMode = False
UserControl.PropertyChanged "OverTypeMode"
End Property

Private Sub CreateTextBox()
If TextBoxHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropRightToLeft = True Then dwExStyle = WS_EX_RTLREADING Or WS_EX_LEFTSCROLLBAR
Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, PropBorderStyle)
Select Case PropAlignment
    Case vbLeftJustify
        dwStyle = dwStyle Or ES_LEFT
    Case vbCenter
        dwStyle = dwStyle Or ES_CENTER
    Case vbRightJustify
        dwStyle = dwStyle Or ES_RIGHT
End Select
If PropAllowOnlyNumbers = True Then dwStyle = dwStyle Or ES_NUMBER
If PropLocked = True Then dwStyle = dwStyle Or ES_READONLY
If PropHideSelection = False Then dwStyle = dwStyle Or ES_NOHIDESEL
If PropUseSystemPasswordChar = True Then dwStyle = dwStyle Or ES_PASSWORD
If PropMultiLine = True Then
    dwStyle = dwStyle Or ES_MULTILINE
    Select Case PropScrollBars
        Case vbSBNone
            dwStyle = dwStyle Or ES_AUTOVSCROLL
        Case vbHorizontal
            dwStyle = dwStyle Or WS_HSCROLL Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
        Case vbVertical
            dwStyle = dwStyle Or WS_VSCROLL Or ES_AUTOVSCROLL
        Case vbBoth
            dwStyle = dwStyle Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
    End Select
Else
    dwStyle = dwStyle Or ES_AUTOHSCROLL
End If
Select Case PropCharacterCasing
    Case TxtCharacterCasingUpper
        dwStyle = dwStyle Or ES_UPPERCASE
    Case TxtCharacterCasingLower
        dwStyle = dwStyle Or ES_LOWERCASE
End Select
If PropNetAddressValidator = True And ComCtlsSupportLevel() >= 2 Then
    If InitNetworkAddressControl() <> 0 Then TextBoxHandle = CreateWindowEx(dwExStyle, StrPtr("msctls_netaddress"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
End If
If TextBoxHandle = 0 Then TextBoxHandle = CreateWindowEx(dwExStyle, StrPtr("Edit"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If TextBoxHandle <> 0 Then
    If PropPasswordChar <> 0 And PropUseSystemPasswordChar = False Then SendMessage TextBoxHandle, EM_SETPASSWORDCHAR, PropPasswordChar, ByVal 0&
    SendMessage TextBoxHandle, EM_SETLIMITTEXT, PropMaxLength, ByVal 0&
    SendMessage TextBoxHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If Not PropCueBanner = vbNullString Then Me.CueBanner = PropCueBanner
If PropNetAddressValidator = True Then Me.NetAddressType = PropNetAddressType
If TextBoxDesignMode = False Then
    If TextBoxHandle <> 0 Then Call ComCtlsSetSubclass(TextBoxHandle, Me, 1)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
    If TextBoxHandle <> 0 Then Call ComCtlsCreateIMC(TextBoxHandle, TextBoxIMCHandle)
End If
End Sub

Private Sub ReCreateTextBox()
If TextBoxDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Dim SelStart As Long, SelEnd As Long
    Dim ScrollPosHorz As Integer, ScrollPosVert As Integer
    If TextBoxHandle <> 0 Then
        SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
        If PropMultiLine = True And PropScrollBars <> vbSBNone Then
            If PropScrollBars = vbHorizontal Or PropScrollBars = vbBoth Then
                ScrollPosHorz = CUIntToInt(GetScrollPos(TextBoxHandle, SB_HORZ) And &HFFFF&)
            End If
            If PropScrollBars = vbVertical Or PropScrollBars = vbBoth Then
                ScrollPosVert = CUIntToInt(GetScrollPos(TextBoxHandle, SB_VERT) And &HFFFF&)
            End If
        End If
        Dim Buffer As String
        Buffer = String$(SendMessage(TextBoxHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
        SendMessage TextBoxHandle, WM_GETTEXT, Len(Buffer) + 1, ByVal StrPtr(Buffer)
        PropText = Buffer
    End If
    Call DestroyTextBox
    Call CreateTextBox
    Call UserControl_Resize
    If TextBoxHandle <> 0 Then
        SendMessage TextBoxHandle, EM_SETSEL, SelStart, ByVal SelEnd
        If ScrollPosHorz > 0 Then SendMessage TextBoxHandle, WM_HSCROLL, MakeDWord(SB_THUMBPOSITION, ScrollPosHorz), ByVal 0&
        If ScrollPosVert > 0 Then SendMessage TextBoxHandle, WM_VSCROLL, MakeDWord(SB_THUMBPOSITION, ScrollPosVert), ByVal 0&
    End If
    If Locked = True Then LockWindowUpdate 0
    Me.Refresh
Else
    Call DestroyTextBox
    Call CreateTextBox
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyTextBox()
If TextBoxHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(TextBoxHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
Call ComCtlsDestroyIMC(TextBoxHandle, TextBoxIMCHandle)
ShowWindow TextBoxHandle, SW_HIDE
SetParent TextBoxHandle, 0
DestroyWindow TextBoxHandle
TextBoxHandle = 0
If TextBoxFontHandle <> 0 Then
    DeleteObject TextBoxFontHandle
    TextBoxFontHandle = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub Copy()
Attribute Copy.VB_Description = "Method to copy the current selection to the clipboard."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_COPY, 0, ByVal 0&
End Sub

Public Sub Cut()
Attribute Cut.VB_Description = "Method to delete (cut) the current selection and copy the deleted text to the clipboard."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_CUT, 0, ByVal 0&
End Sub

Public Sub Paste()
Attribute Paste.VB_Description = "Method to copy the current content of the clipboard at the current caret position."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_PASTE, 0, ByVal 0&
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Method to delete (clear) the current selection."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, WM_CLEAR, 0, ByVal 0&
End Sub

Public Sub Undo()
Attribute Undo.VB_Description = "Undoes the last operation, if any."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_UNDO, 0, ByVal 0&
End Sub

Public Function CanUndo() As Boolean
Attribute CanUndo.VB_Description = "Determines whether there are any actions in the undo queue."
If TextBoxHandle <> 0 Then CanUndo = CBool(SendMessage(TextBoxHandle, EM_CANUNDO, 0, ByVal 0&) <> 0)
End Function

Public Sub ResetUndoFlag()
Attribute ResetUndoFlag.VB_Description = "Resets the undo flag."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
End Sub

Public Property Get Modified() As Boolean
Attribute Modified.VB_Description = "Setting the text property will reset this property to false. Any typing in the control will set the property to true."
Attribute Modified.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then Modified = CBool(SendMessage(TextBoxHandle, EM_GETMODIFY, 0, ByVal 0&) <> 0)
End Property

Public Property Let Modified(ByVal Value As Boolean)
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMODIFY, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Property Get TextLength() As Long
Attribute TextLength.VB_Description = "Returns the length of the text."
Attribute TextLength.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then
    TextLength = SendMessage(TextBoxHandle, WM_GETTEXTLENGTH, 0, ByVal 0&)
Else
    TextLength = Len(PropText)
End If
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute SelStart.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal 0&
End Property

Public Property Let SelStart(ByVal Value As Long)
If TextBoxHandle <> 0 Then
    If Value >= 0 Then
        SendMessage TextBoxHandle, EM_SETSEL, Value, ByVal Value
        SendMessage TextBoxHandle, EM_SCROLLCARET, 0, ByVal 0&
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    SelLength = SelEnd - SelStart
End If
End Property

Public Property Let SelLength(ByVal Value As Long)
If TextBoxHandle <> 0 Then
    If Value >= 0 Then
        Dim SelStart As Long
        SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal 0&
        SendMessage TextBoxHandle, EM_SETSEL, SelStart, ByVal SelStart + Value
        SendMessage TextBoxHandle, EM_SCROLLCARET, 0, ByVal 0&
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    On Error Resume Next
    SelText = Mid$(Me.Text, SelStart + 1, (SelEnd - SelStart))
    On Error GoTo 0
End If
End Property

Public Property Let SelText(ByVal Value As String)
If TextBoxHandle <> 0 Then
    If StrPtr(Value) = 0 Then Value = ""
    SendMessage TextBoxHandle, EM_REPLACESEL, 1, ByVal StrPtr(Value)
End If
End Property

Public Function GetLine(ByVal LineNumber As Long) As String
Attribute GetLine.VB_Description = "Retrieves the text of the specified line. A value of 0 indicates that the text of the current line number (the line that contains the caret) will be retrieved."
If LineNumber < 0 Then Err.Raise 380
If TextBoxHandle <> 0 Then
    Dim FirstCharPos As Long, Length As Long
    FirstCharPos = SendMessage(TextBoxHandle, EM_LINEINDEX, LineNumber - 1, ByVal 0&)
    If FirstCharPos > -1 Then
        Length = SendMessage(TextBoxHandle, EM_LINELENGTH, FirstCharPos, ByVal 0&)
        If Length > 0 Then
            Dim Buffer As String
            Buffer = ChrW(Length) & String$(Length - 1, vbNullChar)
            If LineNumber > 0 Then
                If SendMessage(TextBoxHandle, EM_GETLINE, LineNumber - 1, ByVal StrPtr(Buffer)) > 0 Then GetLine = Buffer
            Else
                If SendMessage(TextBoxHandle, EM_GETLINE, SendMessage(TextBoxHandle, EM_LINEFROMCHAR, FirstCharPos, ByVal 0&), ByVal StrPtr(Buffer)) > 0 Then GetLine = Buffer
            End If
        End If
    Else
        Err.Raise 380
    End If
End If
End Function

Public Function GetLineCount() As Long
Attribute GetLineCount.VB_Description = "Gets the number of lines."
If TextBoxHandle <> 0 Then GetLineCount = SendMessage(TextBoxHandle, EM_GETLINECOUNT, 0, ByVal 0&)
End Function

Public Sub ScrollToLine(ByVal LineNumber As Long)
Attribute ScrollToLine.VB_Description = "Scrolls to ensure the specified line is visible."
If LineNumber < 0 Then Err.Raise 380
If TextBoxHandle <> 0 Then
    If SendMessage(TextBoxHandle, EM_LINEINDEX, LineNumber - 1, ByVal 0&) > -1 Then
        Dim LineIndex As Long
        LineIndex = SendMessage(TextBoxHandle, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&)
        SendMessage TextBoxHandle, EM_LINESCROLL, 0, ByVal CLng((LineNumber - 1) - LineIndex)
    Else
        Err.Raise 380
    End If
End If
End Sub

Public Sub ScrollToCaret()
Attribute ScrollToCaret.VB_Description = "Scrolls the caret into view."
If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SCROLLCARET, 0, ByVal 0&
End Sub

Public Function CharFromPos(ByVal X As Single, ByVal Y As Single) As Long
Attribute CharFromPos.VB_Description = "Returns the character index closest to a specified point."
Dim P As POINTAPI
P.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
P.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
If TextBoxHandle <> 0 Then CharFromPos = CIntToUInt(LoWord(SendMessage(TextBoxHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(P.X, P.Y))))
End Function

Public Function GetLineFromChar(ByVal CharIndex As Long) As Long
Attribute GetLineFromChar.VB_Description = "Gets the line number that contains the specified character index. A character index of -1 retrieves either the current line or the beginning of the current selection."
If CharIndex < -1 Then Err.Raise 380
If TextBoxHandle <> 0 Then GetLineFromChar = SendMessage(TextBoxHandle, EM_LINEFROMCHAR, CharIndex, ByVal 0&) + 1
End Function

Public Function ShowBalloonTip(ByVal Text As String, Optional ByVal Title As String, Optional ByVal Icon As TxtIconConstants) As Boolean
Attribute ShowBalloonTip.VB_Description = "Displays a balloon tip. Requires comctl32.dll version 6.0 or higher."
If TextBoxHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim EDITBT As EDITBALLOONTIP
    With EDITBT
    .cbStruct = LenB(EDITBT)
    .pszText = StrPtr(Text)
    .pszTitle = StrPtr(Title)
    Select Case Icon
        Case TxtIconNone, TxtIconInfo, TxtIconWarning, TxtIconError
            .iIcon = Icon
        Case Else
            Err.Raise 380
    End Select
    If GetFocus() <> TextBoxHandle Then SetFocusAPI UserControl.hWnd
    ShowBalloonTip = CBool(SendMessage(TextBoxHandle, EM_SHOWBALLOONTIP, 0, ByVal VarPtr(EDITBT)) <> 0)
    End With
End If
End Function

Public Function HideBalloonTip() As Boolean
Attribute HideBalloonTip.VB_Description = "Hides any associated balloon tip. Requires comctl32.dll version 6.0 or higher."
If TextBoxHandle <> 0 And ComCtlsSupportLevel() >= 1 Then HideBalloonTip = CBool(SendMessage(TextBoxHandle, EM_HIDEBALLOONTIP, 0, ByVal 0&) <> 0)
End Function

Public Property Get LeftMargin() As Single
Attribute LeftMargin.VB_Description = "Returns/sets the widths of the left margin."
Attribute LeftMargin.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then LeftMargin = UserControl.ScaleX(LoWord(SendMessage(TextBoxHandle, EM_GETMARGINS, 0, ByVal 0&)), vbPixels, vbContainerSize)
End Property

Public Property Let LeftMargin(ByVal Value As Single)
If Value = EC_USEFONTINFO Or Value = -1 Then
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMARGINS, EC_LEFTMARGIN, ByVal EC_USEFONTINFO
Else
    If Value < 0 Then Err.Raise 380
    Dim IntValue As Integer
    IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMARGINS, EC_LEFTMARGIN, ByVal MakeDWord(IntValue, 0)
End If
End Property

Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Returns/sets the widths of the right margin."
Attribute RightMargin.VB_MemberFlags = "400"
If TextBoxHandle <> 0 Then RightMargin = UserControl.ScaleX(HiWord(SendMessage(TextBoxHandle, EM_GETMARGINS, 0, ByVal 0&)), vbPixels, vbContainerSize)
End Property

Public Property Let RightMargin(ByVal Value As Single)
If Value = EC_USEFONTINFO Or Value = -1 Then
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMARGINS, EC_RIGHTMARGIN, ByVal EC_USEFONTINFO
Else
    If Value < 0 Then Err.Raise 380
    Dim IntValue As Integer
    IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    If TextBoxHandle <> 0 Then SendMessage TextBoxHandle, EM_SETMARGINS, EC_RIGHTMARGIN, ByVal MakeDWord(0, IntValue)
End If
End Property

Public Sub ValidateNetAddress()
Attribute ValidateNetAddress.VB_Description = "Validate a network address against a preset network address type mask. Requires comctl32.dll version 6.1 or higher."
TextBoxNetAddressFormat = TxtNetAddressFormatNone
TextBoxNetAddressString = vbNullString
TextBoxNetAddressPortNumber = 0
TextBoxNetAddressPrefixLength = 0
If TextBoxHandle <> 0 And PropNetAddressValidator = True Then
    If ComCtlsSupportLevel() >= 2 Then
        Dim NCADDR As NC_ADDRESS, NETADDRINFO_UNSPECIFIED As NET_ADDRESS_INFO_UNSPECIFIED, ErrVal As Long
        NCADDR.pAddrInfo = VarPtr(NETADDRINFO_UNSPECIFIED)
        ErrVal = SendMessage(TextBoxHandle, NCM_GETADDRESS, 0, ByVal VarPtr(NCADDR))
        Const ERROR_SUCCESS As Long = &H0, S_FALSE As Long = &H1, ERROR_INSUFFICIENT_BUFFER As Long = &H7A, ERROR_INVALID_PARAMETER As Long = &H57, E_INVALIDARG As Long = &H80070057
        Select Case ErrVal
            Case ERROR_SUCCESS
                TextBoxNetAddressFormat = NETADDRINFO_UNSPECIFIED.Format
                TextBoxNetAddressPortNumber = NCADDR.PortNumber
                TextBoxNetAddressPrefixLength = NCADDR.PrefixLength
                Select Case NETADDRINFO_UNSPECIFIED.Format
                    Case NET_ADDRESS_FORMAT_UNSPECIFIED
                        Err.Raise Number:=380, Description:="The network address format is not provided."
                    Case NET_ADDRESS_DNS_NAME
                        Dim NETADDRINFO_DNSNAME As NET_ADDRESS_INFO_DNS_NAME
                        CopyMemory ByVal VarPtr(NETADDRINFO_DNSNAME), NETADDRINFO_UNSPECIFIED.Data(0), LenB(NETADDRINFO_DNSNAME)
                        TextBoxNetAddressString = Left$(NETADDRINFO_DNSNAME.Address(), InStr(NETADDRINFO_DNSNAME.Address(), vbNullChar) - 1)
                    Case NET_ADDRESS_IPV4
                        Dim NETADDRINFO_IPV4 As NET_ADDRESS_INFO_IPV4
                        CopyMemory ByVal VarPtr(NETADDRINFO_IPV4), NETADDRINFO_UNSPECIFIED.Data(0), LenB(NETADDRINFO_IPV4)
                        With NETADDRINFO_IPV4
                        TextBoxNetAddressString = HiByte(HiWord(.sin_addr)) & "." & LoByte(HiWord(.sin_addr)) & "." & HiByte(LoWord(.sin_addr)) & "." & LoByte(LoWord(.sin_addr))
                        End With
                    Case NET_ADDRESS_IPV6
                        Dim NETADDRINFO_IPV6 As NET_ADDRESS_INFO_IPV6, Buffer As String, Temp As String, i As Long
                        CopyMemory ByVal VarPtr(NETADDRINFO_IPV6), NETADDRINFO_UNSPECIFIED.Data(0), LenB(NETADDRINFO_IPV6)
                        With NETADDRINFO_IPV6
                        For i = 1 To 8
                            Temp = Format(Hex(LoByte(.sin6_addr(i - 1))), "00") & Format(Hex(HiByte(.sin6_addr(i - 1))), "00")
                            Do While Left$(Temp, 1) = "0"
                                If Len(Temp) = 1 Then Exit Do
                                Temp = Mid$(Temp, 2)
                            Loop
                            Buffer = Buffer & Temp & ":"
                        Next i
                        TextBoxNetAddressString = Mid$(Buffer, 1, Len(Buffer) - 1) ' Uncompressed IPv6 format
                        End With
                    Case Else
                        Err.Raise Number:=380, Description:="The network address format is unspecified."
                End Select
            Case S_FALSE
                Err.Raise Number:=380, Description:="There is no network address string to validate."
            Case ERROR_INSUFFICIENT_BUFFER
                Err.Raise Number:=ERROR_INSUFFICIENT_BUFFER, Description:="The out buffer is too small to hold the parsed network address."
            Case ERROR_INVALID_PARAMETER
                Err.Raise Number:=ERROR_INVALID_PARAMETER, Description:="The network address string is not of any type specified."
            Case E_INVALIDARG
                Err.Raise Number:=E_INVALIDARG, Description:="The network address string is invalid."
            Case Else
                Err.Raise Number:=ErrVal, Description:="Unexpected error."
        End Select
    Else
        Err.Raise Number:=5, Description:="To use this functionality, you must provide a manifest specifying comctl32.dll version 6.1 or higher."
    End If
Else
    Err.Raise Number:=5, Description:="Procedure call can't be carried out as property NetAddressValidator is False."
End If
End Sub

Public Sub ShowNetAddressErrorTip()
Attribute ShowNetAddressErrorTip.VB_Description = "Display an error ballon tip when an network address string is invalid. Requires comctl32.dll version 6.1 or higher."
If TextBoxHandle <> 0 And PropNetAddressValidator = True And ComCtlsSupportLevel() >= 2 Then
    If GetFocus() <> TextBoxHandle Then SetFocusAPI UserControl.hWnd
    SendMessage TextBoxHandle, NCM_DISPLAYERRORTIP, 0, ByVal 0&
End If
End Sub

Public Property Get NetAddressFormat() As TxtNetAddressFormatConstants
Attribute NetAddressFormat.VB_Description = "Returns the network address format from the latest validation."
Attribute NetAddressFormat.VB_MemberFlags = "400"
NetAddressFormat = TextBoxNetAddressFormat
End Property

Public Property Get NetAddressString() As String
Attribute NetAddressString.VB_Description = "Returns the network address string from the latest validation."
Attribute NetAddressString.VB_MemberFlags = "400"
NetAddressString = TextBoxNetAddressString
End Property

Public Property Get NetAddressPortNumber() As Integer
Attribute NetAddressPortNumber.VB_Description = "Returns the network address port number from the latest validation."
Attribute NetAddressPortNumber.VB_MemberFlags = "400"
NetAddressPortNumber = TextBoxNetAddressPortNumber
End Property

Public Property Get NetAddressPrefixLength() As Byte
Attribute NetAddressPrefixLength.VB_Description = "Returns the network address prefix length from the latest validation."
Attribute NetAddressPrefixLength.VB_MemberFlags = "400"
NetAddressPrefixLength = TextBoxNetAddressPrefixLength
End Property

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If PropOLEDragMode = vbOLEDragAutomatic Then
                Dim P1 As POINTAPI
                Dim CharPos As Long, CaretPos As Long
                Dim SelStart As Long, SelEnd As Long
                GetCursorPos P1
                ScreenToClient TextBoxHandle, P1
                CharPos = CIntToUInt(LoWord(SendMessage(TextBoxHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(P1.X, P1.Y))))
                CaretPos = SendMessage(TextBoxHandle, EM_POSFROMCHAR, CharPos, ByVal 0&)
                SendMessage TextBoxHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                TextBoxAutoDragInSel = CBool(CharPos >= SelStart And CharPos <= SelEnd And CaretPos > -1 And (SelEnd - SelStart) > 0)
                If TextBoxAutoDragInSel = True Then
                    SetCursor LoadCursor(0, MousePointerID(vbArrow))
                    WindowProcControl = 1
                    Exit Function
                End If
            Else
                TextBoxAutoDragInSel = False
            End If
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
    Case WM_LBUTTONDOWN
        If PropOLEDragMode = vbOLEDragAutomatic And TextBoxAutoDragInSel = True Then
            If GetFocus() <> hWnd Then SetFocusAPI UserControl.hWnd ' UCNoSetFocusFwd not applicable
            Dim P2 As POINTAPI, P3 As POINTAPI
            P2.X = Get_X_lParam(lParam)
            P2.Y = Get_Y_lParam(lParam)
            P3.X = P2.X
            P3.Y = P2.Y
            ClientToScreen TextBoxHandle, P3
            RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(P2.X, vbPixels, vbTwips), UserControl.ScaleY(P2.Y, vbPixels, vbTwips))
            If DragDetect(TextBoxHandle, CUIntToInt(P3.X And &HFFFF&), CUIntToInt(P3.Y And &HFFFF&)) <> 0 Then
                TextBoxIsClick = False
                Me.OLEDrag
                WindowProcControl = 0
            Else
                WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
                ReleaseCapture
                RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(P2.X, vbPixels, vbTwips), UserControl.ScaleY(P2.Y, vbPixels, vbTwips))
            End If
            Exit Function
        Else
            If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
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
            If KeyCode = vbKeyInsert And PropAllowOverType = True Then
                If wMsg = WM_KEYDOWN Then PropOverTypeMode = Not PropOverTypeMode
            End If
            TextBoxCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If TextBoxCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(TextBoxCharCodeCache And &HFFFF&)
            TextBoxCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        If (wParam And &HFFFF&) <> 0 And KeyChar = 0 Then
            Exit Function
        Else
            wParam = CIntToUInt(KeyChar)
        End If
        If PropAllowOverType = True And PropOverTypeMode = True Then
            If wParam >= 32 Then ' 0 to 31 are non-printable
                If Me.SelLength = 0 Then
                    Dim FirstCharPos As Long, Length As Long
                    FirstCharPos = SendMessage(TextBoxHandle, EM_LINEINDEX, -1, ByVal 0&)
                    If FirstCharPos > -1 Then
                        Length = SendMessage(TextBoxHandle, EM_LINELENGTH, FirstCharPos, ByVal 0&)
                        If Length > 0 Then
                            If Me.SelStart < (FirstCharPos + Length) Then
                                Me.SelLength = 1
                                Me.SelText = vbNullString
                            End If
                        End If
                    End If
                End If
            End If
        End If
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
    Case WM_INPUTLANGCHANGE
        Call ComCtlsSetIMEMode(hWnd, TextBoxIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call ComCtlsSetIMEMode(hWnd, TextBoxIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_VSCROLL, WM_HSCROLL
        ' The notification codes EN_HSCROLL and EN_VSCROLL are not sent when clicking the scroll bar thumb itself.
        If LoWord(wParam) = SB_THUMBTRACK Then RaiseEvent Scroll
    Case WM_CONTEXTMENU
        If wParam = TextBoxHandle Then
            Dim P4 As POINTAPI, Handled As Boolean
            P4.X = Get_X_lParam(lParam)
            P4.Y = Get_Y_lParam(lParam)
            If P4.X = -1 And P4.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(Handled, -1, -1)
            Else
                ScreenToClient TextBoxHandle, P4
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P4.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P4.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
    Case WM_SETTEXT
        If TextBoxChangeFrozen = False And PropMultiLine = True Then
            ' According to MSDN:
            ' The EN_CHANGE notification code is not sent when the ES_MULTILINE style is used and the text is sent through WM_SETTEXT.
            Dim Buffer(0 To 1) As String
            Buffer(0) = String$(SendMessage(hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
            SendMessage hWnd, WM_GETTEXT, Len(Buffer(0)) + 1, ByVal StrPtr(Buffer(0))
            If lParam <> 0 Then
                Buffer(1) = String$(lstrlen(lParam), vbNullChar)
                CopyMemory ByVal StrPtr(Buffer(1)), ByVal lParam, LenB(Buffer(1))
            End If
            If Buffer(0) <> Buffer(1) Then
                WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
                UserControl.PropertyChanged "Text"
                On Error Resume Next
                UserControl.Extender.DataChanged = True
                On Error GoTo 0
                RaiseEvent Change
                Exit Function
            End If
        End If
    Case WM_PASTE
        If PropAllowOnlyNumbers = True Then
            If ComCtlsSupportLevel() <= 1 Then
                Dim Text As String
                Text = GetClipboardText()
                If Not Text = vbNullString Then
                    Dim i As Long, InvalidText As Boolean
                    For i = 1 To Len(Text)
                        If InStr("0123456789", Mid$(Text, i, 1)) = 0 Then
                            InvalidText = True
                            Exit For
                        End If
                    Next i
                    If InvalidText = True Then
                        VBA.Interaction.Beep
                        Exit Function
                    End If
                End If
            End If
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
                TextBoxIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                TextBoxIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                TextBoxIsClick = True
            Case WM_MOUSEMOVE
                If TextBoxMouseOver = False And PropMouseTrack = True Then
                    TextBoxMouseOver = True
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
                If TextBoxIsClick = True Then
                    TextBoxIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If TextBoxMouseOver = True Then
            TextBoxMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        Select Case HiWord(wParam)
            Case EN_CHANGE
                If TextBoxChangeFrozen = False Then
                    UserControl.PropertyChanged "Text"
                    On Error Resume Next
                    UserControl.Extender.DataChanged = True
                    On Error GoTo 0
                    RaiseEvent Change
                End If
            Case EN_MAXTEXT
                RaiseEvent MaxText
            Case EN_HSCROLL, EN_VSCROLL
                ' This notification code is also sent when a keyboard event causes a change in the view area.
                RaiseEvent Scroll
        End Select
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI TextBoxHandle
End Function
