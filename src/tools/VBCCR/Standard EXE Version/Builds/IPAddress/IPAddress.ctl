VERSION 5.00
Begin VB.UserControl IPAddress 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   MousePointer    =   3  'I-Beam
   PropertyPages   =   "IPAddress.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "IPAddress.ctx":004A
End
Attribute VB_Name = "IPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const ImplementThemedBorder = True

#If False Then
Private IpaAutoSelectNone, IpaAutoSelectFirst, IpaAutoSelectSecond, IpaAutoSelectThird, IpaAutoSelectFourth, IpaAutoSelectBlank
#End If
Public Enum IpaAutoSelectConstants
IpaAutoSelectNone = 0
IpaAutoSelectFirst = 1
IpaAutoSelectSecond = 2
IpaAutoSelectThird = 3
IpaAutoSelectFourth = 4
IpaAutoSelectBlank = 5
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
Private Type TRACKMOUSEEVENTSTRUCT
cbSize As Long
dwFlags As Long
hWndTrack As Long
dwHoverTime As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event SelChange()
Attribute SelChange.VB_Description = "Occurs when the selected item changes."
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
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutW" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long

#If ImplementThemedBorder = True Then

Private Enum UxThemeEditParts
EP_EDITTEXT = 1
EP_CARET = 2
EP_BACKGROUND = 3
EP_PASSWORD = 4
EP_BACKGROUNDWITHBORDER = 5
EP_EDITBORDER_NOSCROLL = 6
EP_EDITBORDER_HSCROLL = 7
EP_EDITBORDER_VSCROLL = 8
EP_EDITBORDER_HVSCROLL = 9
End Enum
Private Enum UxThemeEditBorderNoScrollStates
EPSN_NORMAL = 1
EPSN_HOT = 2
EPSN_FOCUSED = 3
EPSN_DISABLED = 4
End Enum
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal Theme As Long) As Long
Private Declare Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme" (ByVal Theme As Long, iPartId As Long, iStateId As Long) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme" (ByVal hWnd As Long, ByVal hDC As Long, ByRef pRect As RECT) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef pClipRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hRgnClip As Long, ByVal fdwOptions As Long) As Long
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

#End If

Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80, RDW_NOCHILDREN As Long = &H40, RDW_FRAME As Long = &H400
Private Const HWND_DESKTOP As Long = &H0
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const DCX_WINDOW As Long = &H1
Private Const DCX_INTERSECTRGN As Long = &H80
Private Const DCX_USESTYLE As Long = &H10000
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_RTLREADING As Long = &H2000, WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const SW_HIDE As Long = &H0
Private Const TME_LEAVE As Long = &H2, TME_NONCLIENT As Long = &H10
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_ENABLE As Long = &HA
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_STYLECHANGING As Long = &H7C
Private Const WM_STYLECHANGED As Long = &H7D
Private Const WM_COMMAND As Long = &H111
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
Private Const WM_NCMOUSEMOVE As Long = &HA0
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_NCMOUSELEAVE As Long = &H2A2
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_SETFONT As Long = &H30
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_PASTE As Long = &H302
Private Const WM_PAINT As Long = &HF
Private Const WM_NCPAINT As Long = &H85
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_USER As Long = &H400
Private Const IPM_CLEARADDRESS As Long = (WM_USER + 100)
Private Const IPM_SETADDRESS As Long = (WM_USER + 101)
Private Const IPM_GETADDRESS As Long = (WM_USER + 102)
Private Const IPM_SETRANGE As Long = (WM_USER + 103)
Private Const IPM_SETFOCUS As Long = (WM_USER + 104)
Private Const IPM_ISBLANK As Long = (WM_USER + 105)
Private Const EM_SETREADONLY As Long = &HCF, ES_READONLY As Long = &H800
Private Const EM_GETSEL As Long = &HB0
Private Const EM_SETSEL As Long = &HB1
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETLIMITTEXT As Long = EM_LIMITTEXT
Private Const EN_CHANGE As Long = &H300
Private Const ES_LEFT As Long = &H0
Private Const ES_CENTER As Long = &H1
Private Const ES_RIGHT As Long = &H2
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const ES_NUMBER As Long = &H2000
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private IPAddressEditHandle(1 To 4) As Long
Private IPAddressFontHandle As Long
Private IPAddressCharCodeCache As Long
Private IPAddressIsClick As Boolean
Private IPAddressMouseOver(0 To 1) As Boolean
Private IPAddressEditMouseOver(1 To 4) As Boolean
Private IPAddressDesignMode As Boolean
Private IPAddressDotSpacing As Long
Private IPAddressPadding As SIZEAPI
Private IPAddressChangeFrozen As Boolean
Private IPAddressRTLReading(1 To 4) As Boolean
Private IPAddressEnabledVisualStyles As Boolean
Private IPAddressEditFocusHwnd As Long
Private IPAddressSelectedItem As Integer
Private IPAddressMin(1 To 4) As Integer, IPAddressMax(1 To 4) As Integer
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBorderStyle As CCBorderStyleConstants
Private PropText As String
Private PropAutoSelect As IpaAutoSelectConstants
Private PropLocked As Boolean

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
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyTab, vbKeyReturn, vbKeyEscape
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
Call ComCtlsInitCC(ICC_STANDARD_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
IPAddressPadding.CX = 3 * PixelsPerDIP_X()
IPAddressPadding.CY = 1 * PixelsPerDIP_Y()
IPAddressSelectedItem = 1
IPAddressMin(1) = 0
IPAddressMin(2) = 0
IPAddressMin(3) = 0
IPAddressMin(4) = 0
IPAddressMax(1) = 255
IPAddressMax(2) = 255
IPAddressMax(3) = 255
IPAddressMax(4) = 255
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
IPAddressDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = CCBorderStyleSunken ' UserControl.BorderStyle = vbFixedSingle
PropText = vbNullString
PropAutoSelect = IpaAutoSelectFirst
PropLocked = False
Call CreateIPAddress
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
IPAddressDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbWindowBackground)
Me.ForeColor = .ReadProperty("ForeColor", vbWindowText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
Me.BorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSunken)
PropText = .ReadProperty("Text", vbNullString) ' Unicode not necessary
PropAutoSelect = .ReadProperty("AutoSelect", IpaAutoSelectFirst)
PropLocked = .ReadProperty("Locked", False)
End With
Call CreateIPAddress
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbWindowBackground
.WriteProperty "ForeColor", Me.ForeColor, vbWindowText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BorderStyle", Me.BorderStyle, CCBorderStyleSunken
.WriteProperty "Text", PropText, vbNullString ' Unicode not necessary
.WriteProperty "AutoSelect", PropAutoSelect, IpaAutoSelectFirst
.WriteProperty "Locked", PropLocked, False
End With
End Sub

Private Sub UserControl_Paint()
UserControl.Cls
Call DrawDots(UserControl.hDC)
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (IPAddressMouseOver(0) = False And PropMouseTrack = True) Or (IPAddressMouseOver(1) = False And PropMouseTrack = True) Then
    If IPAddressMouseOver(0) = False And PropMouseTrack = True Then IPAddressMouseOver(0) = True
    If IPAddressMouseOver(1) = False And PropMouseTrack = True Then
        IPAddressMouseOver(1) = True
        RaiseEvent MouseEnter
    End If
    Call ComCtlsRequestMouseLeave(hWnd)
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
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
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
Dim X As Long, Y As Long, CX As Long, CY As Long
X = IPAddressPadding.CX
Y = IPAddressPadding.CY
CX = ((UserControl.ScaleWidth - X) - (IPAddressDotSpacing * 3)) \ 4 ' Discard any remainder
CY = UserControl.ScaleHeight - Y
If IPAddressEditHandle(1) <> 0 Then MoveWindow IPAddressEditHandle(1), X, Y, CX, CY, 1
X = X + CX + IPAddressDotSpacing
If IPAddressEditHandle(2) <> 0 Then MoveWindow IPAddressEditHandle(2), X, Y, CX, CY, 1
X = X + CX + IPAddressDotSpacing
If IPAddressEditHandle(3) <> 0 Then MoveWindow IPAddressEditHandle(3), X, Y, CX, CY, 1
X = X + CX + IPAddressDotSpacing
If IPAddressEditHandle(4) <> 0 Then MoveWindow IPAddressEditHandle(4), X, Y, CX, CY, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyIPAddress
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
hWnd = UserControl.hWnd
End Property

Public Property Get hWndEdit(ByVal Index As Integer) As Long
Attribute hWndEdit.VB_Description = "Returns a handle to a control."
If Index > 4 Or Index < 1 Then Err.Raise Number:=35600, Description:="Index out of bounds"
hWndEdit = IPAddressEditHandle(Index)
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
OldFontHandle = IPAddressFontHandle
IPAddressFontHandle = CreateGDIFontFromOLEFont(PropFont)
If IPAddressEditHandle(1) <> 0 Then SendMessage IPAddressEditHandle(1), WM_SETFONT, IPAddressFontHandle, ByVal 1&
If IPAddressEditHandle(2) <> 0 Then SendMessage IPAddressEditHandle(2), WM_SETFONT, IPAddressFontHandle, ByVal 1&
If IPAddressEditHandle(3) <> 0 Then SendMessage IPAddressEditHandle(3), WM_SETFONT, IPAddressFontHandle, ByVal 1&
If IPAddressEditHandle(4) <> 0 Then SendMessage IPAddressEditHandle(4), WM_SETFONT, IPAddressFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    Dim Size As SIZEAPI, hFontOld As Long
    If IPAddressFontHandle <> 0 Then hFontOld = SelectObject(hDCScreen, IPAddressFontHandle)
    GetTextExtentPoint32 hDCScreen, StrPtr("."), 1, Size
    IPAddressDotSpacing = Size.CX
    If hFontOld <> 0 Then SelectObject hDCScreen, hFontOld
    ReleaseDC 0, hDCScreen
End If
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = IPAddressFontHandle
IPAddressFontHandle = CreateGDIFontFromOLEFont(PropFont)
If IPAddressEditHandle(1) <> 0 Then SendMessage IPAddressEditHandle(1), WM_SETFONT, IPAddressFontHandle, ByVal 1&
If IPAddressEditHandle(2) <> 0 Then SendMessage IPAddressEditHandle(2), WM_SETFONT, IPAddressFontHandle, ByVal 1&
If IPAddressEditHandle(3) <> 0 Then SendMessage IPAddressEditHandle(3), WM_SETFONT, IPAddressFontHandle, ByVal 1&
If IPAddressEditHandle(4) <> 0 Then SendMessage IPAddressEditHandle(4), WM_SETFONT, IPAddressFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Dim hDCScreen As Long
hDCScreen = GetDC(0)
If hDCScreen <> 0 Then
    Dim Size As SIZEAPI, hFontOld As Long
    If IPAddressFontHandle <> 0 Then hFontOld = SelectObject(hDCScreen, IPAddressFontHandle)
    GetTextExtentPoint32 hDCScreen, StrPtr("."), 1, Size
    IPAddressDotSpacing = Size.CX
    If hFontOld <> 0 Then SelectObject hDCScreen, hFontOld
    ReleaseDC 0, hDCScreen
End If
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
IPAddressEnabledVisualStyles = EnabledVisualStyles()
If IPAddressEnabledVisualStyles = True Then
    If PropVisualStyles = True Then
        If IPAddressEditHandle(1) <> 0 Then ActivateVisualStyles IPAddressEditHandle(1)
        If IPAddressEditHandle(2) <> 0 Then ActivateVisualStyles IPAddressEditHandle(2)
        If IPAddressEditHandle(3) <> 0 Then ActivateVisualStyles IPAddressEditHandle(3)
        If IPAddressEditHandle(4) <> 0 Then ActivateVisualStyles IPAddressEditHandle(4)
    Else
        If IPAddressEditHandle(1) <> 0 Then RemoveVisualStyles IPAddressEditHandle(1)
        If IPAddressEditHandle(2) <> 0 Then RemoveVisualStyles IPAddressEditHandle(2)
        If IPAddressEditHandle(3) <> 0 Then RemoveVisualStyles IPAddressEditHandle(3)
        If IPAddressEditHandle(4) <> 0 Then RemoveVisualStyles IPAddressEditHandle(4)
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

#If ImplementThemedBorder = True Then

If PropBorderStyle = CCBorderStyleSunken Then
    ' Redraw the border to consider the new back color for the themed border, if any.
    RedrawWindow UserControl.hWnd, 0, 0, RDW_FRAME Or RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_NOCHILDREN
End If

#End If

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
If IPAddressEditHandle(1) <> 0 Then EnableWindow IPAddressEditHandle(1), IIf(Value = True, 1, 0)
If IPAddressEditHandle(2) <> 0 Then EnableWindow IPAddressEditHandle(2), IIf(Value = True, 1, 0)
If IPAddressEditHandle(3) <> 0 Then EnableWindow IPAddressEditHandle(3), IIf(Value = True, 1, 0)
If IPAddressEditHandle(4) <> 0 Then EnableWindow IPAddressEditHandle(4), IIf(Value = True, 1, 0)
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
If IPAddressDesignMode = False Then Call RefreshMousePointer
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
        If IPAddressDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If IPAddressDesignMode = False Then Call RefreshMousePointer
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
If IPAddressEditHandle(1) <> 0 Then Call ComCtlsSetRightToLeft(IPAddressEditHandle(1), dwMask)
If IPAddressEditHandle(2) <> 0 Then Call ComCtlsSetRightToLeft(IPAddressEditHandle(2), dwMask)
If IPAddressEditHandle(3) <> 0 Then Call ComCtlsSetRightToLeft(IPAddressEditHandle(3), dwMask)
If IPAddressEditHandle(4) <> 0 Then Call ComCtlsSetRightToLeft(IPAddressEditHandle(4), dwMask)
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
Call ComCtlsChangeBorderStyle(UserControl.hWnd, PropBorderStyle)
Call UserControl_Resize
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "200"
Dim Length(1 To 4) As Long
If IPAddressEditHandle(1) <> 0 Then Length(1) = SendMessage(IPAddressEditHandle(1), WM_GETTEXTLENGTH, 0, ByVal 0&)
If IPAddressEditHandle(2) <> 0 Then Length(2) = SendMessage(IPAddressEditHandle(2), WM_GETTEXTLENGTH, 0, ByVal 0&)
If IPAddressEditHandle(3) <> 0 Then Length(3) = SendMessage(IPAddressEditHandle(3), WM_GETTEXTLENGTH, 0, ByVal 0&)
If IPAddressEditHandle(4) <> 0 Then Length(4) = SendMessage(IPAddressEditHandle(4), WM_GETTEXTLENGTH, 0, ByVal 0&)
If Length(1) > 0 Or Length(2) > 0 Or Length(3) > 0 Or Length(4) > 0 Then
    Dim Buffer(1 To 4) As String
    If Length(1) > 0 Then
        Buffer(1) = String$(Length(1), vbNullChar)
        SendMessage IPAddressEditHandle(1), WM_GETTEXT, Length(1) + 1, ByVal StrPtr(Buffer(1))
    End If
    If Length(2) > 0 Then
        Buffer(2) = String$(Length(2), vbNullChar)
        SendMessage IPAddressEditHandle(2), WM_GETTEXT, Length(2) + 1, ByVal StrPtr(Buffer(2))
    End If
    If Length(3) > 0 Then
        Buffer(3) = String$(Length(3), vbNullChar)
        SendMessage IPAddressEditHandle(3), WM_GETTEXT, Length(3) + 1, ByVal StrPtr(Buffer(3))
    End If
    If Length(4) > 0 Then
        Buffer(4) = String$(Length(4), vbNullChar)
        SendMessage IPAddressEditHandle(4), WM_GETTEXT, Length(4) + 1, ByVal StrPtr(Buffer(4))
    End If
    Text = Buffer(1) & "." & Buffer(2) & "." & Buffer(3) & "." & Buffer(4)
End If
End Property

Public Property Let Text(ByVal Value As String)
Dim OldText As String
OldText = Me.Text
If Value = vbNullString Then
    IPAddressChangeFrozen = True
    If IPAddressEditHandle(1) <> 0 Then SendMessage IPAddressEditHandle(1), WM_SETTEXT, 0, ByVal 0&
    If IPAddressEditHandle(2) <> 0 Then SendMessage IPAddressEditHandle(2), WM_SETTEXT, 0, ByVal 0&
    If IPAddressEditHandle(3) <> 0 Then SendMessage IPAddressEditHandle(3), WM_SETTEXT, 0, ByVal 0&
    If IPAddressEditHandle(4) <> 0 Then SendMessage IPAddressEditHandle(4), WM_SETTEXT, 0, ByVal 0&
    IPAddressChangeFrozen = False
    If Not OldText = vbNullString Then RaiseEvent Change
Else
    Dim Buffer(0 To 3) As String
    Dim Pos1 As Long, Pos2 As Long, i As Long, j As Long
    Do
        If i > 3 Then i = -1: Exit Do
        Pos1 = InStr(Pos1 + 1, Value, ".")
        If Pos1 > 0 Then
            Buffer(i) = Mid$(Value, Pos2 + 1, Pos1 - Pos2 - 1)
        Else
            Buffer(i) = Mid$(Value, Pos2 + 1)
        End If
        Pos2 = Pos1
        i = i + 1
    Loop Until Pos1 = 0
    If i = 4 Then
        Dim InvalidText As Boolean
        For i = 0 To 3
            If Len(Buffer(i)) > 3 Then
                InvalidText = True
                Exit For
            Else
                For j = 1 To Len(Buffer(i))
                    If InStr("0123456789", Mid$(Buffer(i), j, 1)) = 0 Then
                        InvalidText = True
                        Exit For
                    End If
                Next j
            End If
            If InvalidText = True Then Exit For
        Next i
        If InvalidText = False Then
            Dim LngValue(0 To 3) As Long
            For i = 0 To 3
                On Error Resume Next
                LngValue(i) = CLng(Buffer(i))
                On Error GoTo 0
                If LngValue(i) < IPAddressMin(i + 1) Then
                    Buffer(i) = CStr(IPAddressMin(i + 1))
                ElseIf LngValue(i) > IPAddressMax(i + 1) Then
                    Buffer(i) = CStr(IPAddressMax(i + 1))
                End If
                IPAddressChangeFrozen = True
                If IPAddressEditHandle(i + 1) <> 0 Then SendMessage IPAddressEditHandle(i + 1), WM_SETTEXT, 0, ByVal StrPtr(Buffer(i))
                IPAddressChangeFrozen = False
            Next i
            If Not OldText = Buffer(0) & "." & Buffer(1) & "." & Buffer(2) & "." & Buffer(3) Then RaiseEvent Change
        Else
            If IPAddressDesignMode = True Then
                MsgBox "Invalid property value", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise 380
            End If
        End If
    Else
        If IPAddressDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
PropText = Value
UserControl.PropertyChanged "Text"
End Property

Public Property Get AutoSelect() As IpaAutoSelectConstants
Attribute AutoSelect.VB_Description = "Returns/sets which item will be selected automatically upon keyboard focus."
AutoSelect = PropAutoSelect
End Property

Public Property Let AutoSelect(ByVal Value As IpaAutoSelectConstants)
Select Case Value
    Case IpaAutoSelectNone, IpaAutoSelectFirst, IpaAutoSelectSecond, IpaAutoSelectThird, IpaAutoSelectFourth, IpaAutoSelectBlank
        PropAutoSelect = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "AutoSelect"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents can be edited."
Locked = PropLocked
End Property

Public Property Let Locked(ByVal Value As Boolean)
PropLocked = Value
If IPAddressEditHandle(1) <> 0 Then SendMessage IPAddressEditHandle(1), EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
If IPAddressEditHandle(2) <> 0 Then SendMessage IPAddressEditHandle(2), EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
If IPAddressEditHandle(3) <> 0 Then SendMessage IPAddressEditHandle(3), EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
If IPAddressEditHandle(4) <> 0 Then SendMessage IPAddressEditHandle(4), EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Locked"
End Property

Private Sub CreateIPAddress()
If IPAddressEditHandle(1) <> 0 Or IPAddressEditHandle(2) <> 0 Or IPAddressEditHandle(3) <> 0 Or IPAddressEditHandle(4) <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or ES_CENTER Or ES_AUTOHSCROLL Or ES_NUMBER
If PropRightToLeft = True Then dwExStyle = WS_EX_RTLREADING Or WS_EX_LEFTSCROLLBAR
IPAddressRTLReading(1) = CBool((dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING)
If PropLocked = True Then dwStyle = dwStyle Or ES_READONLY
IPAddressEditHandle(1) = CreateWindowEx(dwExStyle, StrPtr("Edit"), 0, dwStyle, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If IPAddressEditHandle(1) <> 0 Then SendMessage IPAddressEditHandle(1), EM_SETLIMITTEXT, 3, ByVal 0&
IPAddressRTLReading(2) = IPAddressRTLReading(1)
IPAddressEditHandle(2) = CreateWindowEx(dwExStyle, StrPtr("Edit"), 0, dwStyle, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If IPAddressEditHandle(2) <> 0 Then SendMessage IPAddressEditHandle(2), EM_SETLIMITTEXT, 3, ByVal 0&
IPAddressRTLReading(3) = IPAddressRTLReading(1)
IPAddressEditHandle(3) = CreateWindowEx(dwExStyle, StrPtr("Edit"), 0, dwStyle, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If IPAddressEditHandle(3) <> 0 Then SendMessage IPAddressEditHandle(3), EM_SETLIMITTEXT, 3, ByVal 0&
IPAddressRTLReading(4) = IPAddressRTLReading(1)
IPAddressEditHandle(4) = CreateWindowEx(dwExStyle, StrPtr("Edit"), 0, dwStyle, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If IPAddressEditHandle(4) <> 0 Then SendMessage IPAddressEditHandle(4), EM_SETLIMITTEXT, 3, ByVal 0&
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Text = PropText
If IPAddressDesignMode = False Then
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 0)
    If IPAddressEditHandle(1) <> 0 Then Call ComCtlsSetSubclass(IPAddressEditHandle(1), Me, 1)
    If IPAddressEditHandle(2) <> 0 Then Call ComCtlsSetSubclass(IPAddressEditHandle(2), Me, 2)
    If IPAddressEditHandle(3) <> 0 Then Call ComCtlsSetSubclass(IPAddressEditHandle(3), Me, 3)
    If IPAddressEditHandle(4) <> 0 Then Call ComCtlsSetSubclass(IPAddressEditHandle(4), Me, 4)
End If
End Sub

Private Sub DestroyIPAddress()
If IPAddressEditHandle(1) = 0 And IPAddressEditHandle(2) = 0 And IPAddressEditHandle(3) = 0 And IPAddressEditHandle(4) = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(UserControl.hWnd)
Call ComCtlsRemoveSubclass(IPAddressEditHandle(1))
Call ComCtlsRemoveSubclass(IPAddressEditHandle(2))
Call ComCtlsRemoveSubclass(IPAddressEditHandle(3))
Call ComCtlsRemoveSubclass(IPAddressEditHandle(4))
ShowWindow IPAddressEditHandle(1), SW_HIDE
ShowWindow IPAddressEditHandle(2), SW_HIDE
ShowWindow IPAddressEditHandle(3), SW_HIDE
ShowWindow IPAddressEditHandle(4), SW_HIDE
SetParent IPAddressEditHandle(1), 0
SetParent IPAddressEditHandle(2), 0
SetParent IPAddressEditHandle(3), 0
SetParent IPAddressEditHandle(4), 0
DestroyWindow IPAddressEditHandle(1)
DestroyWindow IPAddressEditHandle(2)
DestroyWindow IPAddressEditHandle(3)
DestroyWindow IPAddressEditHandle(4)
IPAddressEditHandle(1) = 0
IPAddressEditHandle(2) = 0
IPAddressEditHandle(3) = 0
IPAddressEditHandle(4) = 0
If IPAddressFontHandle <> 0 Then
    DeleteObject IPAddressFontHandle
    IPAddressFontHandle = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get SelectedItem() As Integer
Attribute SelectedItem.VB_Description = "Returns/sets a reference to the currently selected item."
Attribute SelectedItem.VB_MemberFlags = "400"
SelectedItem = IPAddressSelectedItem
End Property

Public Property Let SelectedItem(ByVal Value As Integer)
If Value > 4 Or Value < 1 Then Err.Raise 380
IPAddressSelectedItem = Value
If IPAddressEditFocusHwnd <> 0 Then
    UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
    SetFocusAPI IPAddressEditHandle(IPAddressSelectedItem)
    SendMessage IPAddressEditHandle(IPAddressSelectedItem), EM_SETSEL, 0, ByVal -1&
End If
End Property

Public Property Get Min(ByVal Item As Integer) As Byte
Attribute Min.VB_Description = "Returns/sets the minimum value that the specified item accepts."
Attribute Min.VB_MemberFlags = "400"
If Item > 4 Or Item < 1 Then Err.Raise 5
Min = IPAddressMin(Item)
End Property

Public Property Let Min(ByVal Item As Integer, ByVal Value As Byte)
If Item > 4 Or Item < 1 Then Err.Raise 5
If Value > IPAddressMax(Item) Then Value = IPAddressMax(Item)
IPAddressMin(Item) = Value
End Property

Public Property Get Max(ByVal Item As Integer) As Byte
Attribute Max.VB_Description = "Returns/sets the maximum value that the specified item accepts."
Attribute Max.VB_MemberFlags = "400"
If Item > 4 Or Item < 1 Then Err.Raise 5
Max = IPAddressMax(Item)
End Property

Public Property Let Max(ByVal Item As Integer, ByVal Value As Byte)
If Item > 4 Or Item < 1 Then Err.Raise 5
If Value < IPAddressMin(Item) Then Value = IPAddressMin(Item)
IPAddressMax(Item) = Value
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets a value which represents the text of all four items."
Attribute Value.VB_MemberFlags = "400"
Dim Length(1 To 4) As Long
If IPAddressEditHandle(1) <> 0 Then Length(1) = SendMessage(IPAddressEditHandle(1), WM_GETTEXTLENGTH, 0, ByVal 0&)
If IPAddressEditHandle(2) <> 0 Then Length(2) = SendMessage(IPAddressEditHandle(2), WM_GETTEXTLENGTH, 0, ByVal 0&)
If IPAddressEditHandle(3) <> 0 Then Length(3) = SendMessage(IPAddressEditHandle(3), WM_GETTEXTLENGTH, 0, ByVal 0&)
If IPAddressEditHandle(4) <> 0 Then Length(4) = SendMessage(IPAddressEditHandle(4), WM_GETTEXTLENGTH, 0, ByVal 0&)
Dim Buffer(1 To 4) As String
If Length(1) > 0 Then
    Buffer(1) = String$(Length(1), vbNullChar)
    SendMessage IPAddressEditHandle(1), WM_GETTEXT, Length(1) + 1, ByVal StrPtr(Buffer(1))
End If
If Length(2) > 0 Then
    Buffer(2) = String$(Length(2), vbNullChar)
    SendMessage IPAddressEditHandle(2), WM_GETTEXT, Length(2) + 1, ByVal StrPtr(Buffer(2))
End If
If Length(3) > 0 Then
    Buffer(3) = String$(Length(3), vbNullChar)
    SendMessage IPAddressEditHandle(3), WM_GETTEXT, Length(3) + 1, ByVal StrPtr(Buffer(3))
End If
If Length(4) > 0 Then
    Buffer(4) = String$(Length(4), vbNullChar)
    SendMessage IPAddressEditHandle(4), WM_GETTEXT, Length(4) + 1, ByVal StrPtr(Buffer(4))
End If
Dim LngValue(1 To 4) As Long, i As Long
For i = 1 To 4
    On Error Resume Next
    LngValue(i) = CLng(Buffer(i))
    On Error GoTo 0
Next i
Value = MakeDWord(MakeWord(LngValue(4) And &HFF&, LngValue(3) And &HFF&), MakeWord(LngValue(2) And &HFF&, LngValue(1) And &HFF&))
End Property

Public Property Let Value(ByVal NewValue As Long)
Dim OldText As String
OldText = Me.Text
Dim IntValue(1 To 4) As Integer
IntValue(1) = HiWord(NewValue)
IntValue(2) = LoByte(IntValue(1))
IntValue(1) = HiByte(IntValue(1))
IntValue(3) = LoWord(NewValue)
IntValue(4) = LoByte(IntValue(3))
IntValue(3) = HiByte(IntValue(3))
IPAddressChangeFrozen = True
Dim Buffer As String, NewText As String
Buffer = CStr(IntValue(1))
NewText = Buffer & "."
If IPAddressEditHandle(1) <> 0 Then SendMessage IPAddressEditHandle(1), WM_SETTEXT, 0, ByVal StrPtr(Buffer)
Buffer = CStr(IntValue(2))
NewText = NewText & Buffer & "."
If IPAddressEditHandle(2) <> 0 Then SendMessage IPAddressEditHandle(2), WM_SETTEXT, 0, ByVal StrPtr(Buffer)
Buffer = CStr(IntValue(3))
NewText = NewText & Buffer & "."
If IPAddressEditHandle(3) <> 0 Then SendMessage IPAddressEditHandle(3), WM_SETTEXT, 0, ByVal StrPtr(Buffer)
Buffer = CStr(IntValue(4))
NewText = NewText & Buffer
If IPAddressEditHandle(4) <> 0 Then SendMessage IPAddressEditHandle(4), WM_SETTEXT, 0, ByVal StrPtr(Buffer)
IPAddressChangeFrozen = False
If Not OldText = NewText Then RaiseEvent Change
End Property

Private Sub DrawDots(ByVal hDC As Long)
Dim hFontOld As Long
If IPAddressFontHandle <> 0 Then hFontOld = SelectObject(hDC, IPAddressFontHandle)
Dim X As Long, Y As Long, CX As Long
X = IPAddressPadding.CX
Y = IPAddressPadding.CY
CX = ((UserControl.ScaleWidth - X) - (IPAddressDotSpacing * 3)) \ 4 ' Discard any remainder
X = X + CX
TextOut hDC, X, Y, StrPtr("."), 1
X = X + IPAddressDotSpacing + CX
TextOut hDC, X, Y, StrPtr("."), 1
X = X + IPAddressDotSpacing + CX
TextOut hDC, X, Y, StrPtr("."), 1
If hFontOld <> 0 Then SelectObject hDC, hFontOld
End Sub

Private Function CheckMinMaxFromWindow(ByVal hWnd As Long) As Boolean
Dim Item As Integer
Item = ItemFromWindow(hWnd)
If Item > 0 Then
    Dim StrValue As String
    StrValue = String(SendMessage(hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage hWnd, WM_GETTEXT, Len(StrValue) + 1, ByVal StrPtr(StrValue)
    If Not StrValue = vbNullString Then
        Dim LngValue As Long
        On Error Resume Next
        LngValue = CLng(StrValue)
        On Error GoTo 0
        If LngValue < IPAddressMin(Item) Then
            StrValue = CStr(IPAddressMin(Item))
            SendMessage hWnd, WM_SETTEXT, 0, ByVal StrPtr(StrValue)
            CheckMinMaxFromWindow = True
        ElseIf LngValue > IPAddressMax(Item) Then
            StrValue = CStr(IPAddressMax(Item))
            SendMessage hWnd, WM_SETTEXT, 0, ByVal StrPtr(StrValue)
            CheckMinMaxFromWindow = True
        End If
    End If
End If
End Function

Private Function GetNonBlankCount() As Long
Dim Count As Long
If IPAddressEditHandle(1) <> 0 Then If SendMessage(IPAddressEditHandle(1), WM_GETTEXTLENGTH, 0, ByVal 0&) > 0 Then Count = Count + 1
If IPAddressEditHandle(2) <> 0 Then If SendMessage(IPAddressEditHandle(2), WM_GETTEXTLENGTH, 0, ByVal 0&) > 0 Then Count = Count + 1
If IPAddressEditHandle(3) <> 0 Then If SendMessage(IPAddressEditHandle(3), WM_GETTEXTLENGTH, 0, ByVal 0&) > 0 Then Count = Count + 1
If IPAddressEditHandle(4) <> 0 Then If SendMessage(IPAddressEditHandle(4), WM_GETTEXTLENGTH, 0, ByVal 0&) > 0 Then Count = Count + 1
GetNonBlankCount = Count
End Function

Private Function GetBlankItem() As Integer
Dim NonBlank(1 To 4) As Boolean
If IPAddressEditHandle(1) <> 0 Then
    If SendMessage(IPAddressEditHandle(1), WM_GETTEXTLENGTH, 0, ByVal 0&) > 0 Then NonBlank(1) = True
End If
If IPAddressEditHandle(2) <> 0 Then
    If SendMessage(IPAddressEditHandle(2), WM_GETTEXTLENGTH, 0, ByVal 0&) > 0 Then NonBlank(2) = True
End If
If IPAddressEditHandle(3) <> 0 Then
    If SendMessage(IPAddressEditHandle(3), WM_GETTEXTLENGTH, 0, ByVal 0&) > 0 Then NonBlank(3) = True
End If
If IPAddressEditHandle(4) <> 0 Then
    If SendMessage(IPAddressEditHandle(4), WM_GETTEXTLENGTH, 0, ByVal 0&) > 0 Then NonBlank(4) = True
End If
If NonBlank(1) = True And NonBlank(2) = True And NonBlank(3) = True And NonBlank(4) = True Then
    ' If all are non-blank then set first item.
    GetBlankItem = 1
Else
    If NonBlank(1) = False Then
        GetBlankItem = 1
    ElseIf NonBlank(2) = False Then
        GetBlankItem = 2
    ElseIf NonBlank(3) = False Then
        GetBlankItem = 3
    ElseIf NonBlank(4) = False Then
        GetBlankItem = 4
    End If
End If
End Function

Private Function ItemFromWindow(ByVal hWnd As Long) As Integer
If hWnd <> 0 Then
    Select Case hWnd
        Case IPAddressEditHandle(1)
            ItemFromWindow = 1
        Case IPAddressEditHandle(2)
            ItemFromWindow = 2
        Case IPAddressEditHandle(3)
            ItemFromWindow = 3
        Case IPAddressEditHandle(4)
            ItemFromWindow = 4
    End Select
End If
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 0
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Case 1 To 4
        ISubclass_Message = WindowProcEdit(hWnd, wMsg, wParam, lParam, dwRefData)
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        Select Case HiWord(wParam)
            Case EN_CHANGE
                If IPAddressChangeFrozen = False Then RaiseEvent Change
        End Select
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
                WindowProcUserControl = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcUserControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_PRINTCLIENT
        SendMessage hWnd, WM_PAINT, wParam, ByVal 0&
        Call DrawDots(wParam)
        Dim WndRect1 As RECT, P As POINTAPI, i As Long
        For i = 1 To 4
            GetWindowRect IPAddressEditHandle(i), WndRect1
            MapWindowPoints HWND_DESKTOP, UserControl.hWnd, WndRect1, 2
            SetViewportOrgEx wParam, WndRect1.Left, WndRect1.Top, P
            SendMessage IPAddressEditHandle(i), WM_PAINT, wParam, ByVal 0&
            SetViewportOrgEx wParam, P.X, P.Y, P
        Next i
        WindowProcUserControl = 0
        Exit Function
    
    #If ImplementThemedBorder = True Then
    
    Case WM_THEMECHANGED, WM_STYLECHANGED, WM_ENABLE
        If wMsg = WM_THEMECHANGED Then IPAddressEnabledVisualStyles = EnabledVisualStyles()
        If PropBorderStyle = CCBorderStyleSunken And PropVisualStyles = True Then
            If IPAddressEnabledVisualStyles = True Then SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
        End If
    Case WM_NCPAINT
        If PropBorderStyle = CCBorderStyleSunken And PropVisualStyles = True And IPAddressEnabledVisualStyles = True Then
            Dim Theme As Long
            Theme = OpenThemeData(hWnd, StrPtr("Edit"))
            If Theme <> 0 Then
                Dim hDC As Long
                If wParam = 1 Then ' Alias for entire window
                    hDC = GetWindowDC(hWnd)
                Else
                    hDC = GetDCEx(hWnd, wParam, DCX_WINDOW Or DCX_INTERSECTRGN Or DCX_USESTYLE)
                End If
                If hDC <> 0 Then
                    Dim BorderX As Long, BorderY As Long
                    Dim RC1 As RECT, RC2 As RECT, WndRect2 As RECT
                    Const SM_CXEDGE As Long = 45
                    Const SM_CYEDGE As Long = 46
                    BorderX = GetSystemMetrics(SM_CXEDGE)
                    BorderY = GetSystemMetrics(SM_CYEDGE)
                    GetWindowRect hWnd, WndRect2
                    With UserControl
                    SetRect RC1, BorderX, BorderY, (WndRect2.Right - WndRect2.Left) - BorderX, (WndRect2.Bottom - WndRect2.Top) - BorderY
                    SetRect RC2, 0, 0, (WndRect2.Right - WndRect2.Left), (WndRect2.Bottom - WndRect2.Top)
                    End With
                    ExcludeClipRect hDC, RC1.Left, RC1.Top, RC1.Right, RC1.Bottom
                    Dim EditPart As Long, EditState As Long
                    EditPart = EP_EDITBORDER_NOSCROLL
                    Dim Brush As Long
                    If Me.Enabled = False Then
                        EditState = EPSN_DISABLED
                        Brush = CreateSolidBrush(WinColor(vbButtonFace))
                    Else
                        If IPAddressEditFocusHwnd <> 0 Then
                            EditState = EPSN_FOCUSED
                        Else
                            EditState = EPSN_NORMAL
                        End If
                        Brush = CreateSolidBrush(WinColor(Me.BackColor))
                    End If
                    FillRect hDC, RC2, Brush
                    DeleteObject Brush
                    If IsThemeBackgroundPartiallyTransparent(Theme, EditPart, EditState) <> 0 Then DrawThemeParentBackground hWnd, hDC, RC2
                    DrawThemeBackground Theme, hDC, EditPart, EditState, RC2, RC2
                    ReleaseDC hWnd, hDC
                End If
                CloseThemeData Theme
                WindowProcUserControl = 0
                Exit Function
            End If
        End If
    
    #End If
    
    ' Compatibility for the SysIPAddress32 messages
    
    Case IPM_CLEARADDRESS
        Me.Text = vbNullString
        Exit Function
    Case IPM_SETADDRESS
        Me.Value = lParam
        Exit Function
    Case IPM_GETADDRESS
        Dim LngValue As Long
        LngValue = Me.Value
        If lParam <> 0 Then CopyMemory ByVal lParam, ByVal VarPtr(LngValue), 4
        WindowProcUserControl = GetNonBlankCount()
        Exit Function
    Case IPM_SETRANGE
        Select Case wParam
            Case 0 To 3
                Dim IntValue As Integer
                IntValue = LoWord(lParam)
                Me.Min(wParam + 1) = LoByte(IntValue)
                Me.Max(wParam + 1) = HiByte(IntValue)
                WindowProcUserControl = 1
            Case Else
                WindowProcUserControl = 0
        End Select
        Exit Function
    Case IPM_SETFOCUS
        Dim Item As Integer
        Select Case wParam
            Case Is > 3
                Item = GetBlankItem()
            Case 3
                Item = 4
            Case 2
                Item = 3
            Case 1
                Item = 2
            Case 0
                Item = 1
        End Select
        If Item > 0 Then
            If IPAddressEditHandle(Item) <> 0 Then
                UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
                SetFocusAPI IPAddressEditHandle(Item)
                SendMessage IPAddressEditHandle(Item), EM_SETSEL, 0, ByVal -1&
            End If
        End If
        Exit Function
    Case IPM_ISBLANK
        If Me.Text = vbNullString Then WindowProcUserControl = 1 Else WindowProcUserControl = 0
        Exit Function
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_SETFOCUS
        If UCNoSetFocusFwd = False Then
            Select Case PropAutoSelect
                Case IpaAutoSelectNone
                    If IPAddressEditHandle(IPAddressSelectedItem) <> 0 Then
                        SetFocusAPI IPAddressEditHandle(IPAddressSelectedItem)
                        SendMessage IPAddressEditHandle(IPAddressSelectedItem), EM_SETSEL, 0, ByVal -1&
                    End If
                Case IpaAutoSelectFirst
                    If IPAddressEditHandle(1) <> 0 Then
                        SetFocusAPI IPAddressEditHandle(1)
                        SendMessage IPAddressEditHandle(1), EM_SETSEL, 0, ByVal -1&
                    End If
                Case IpaAutoSelectSecond
                    If IPAddressEditHandle(2) <> 0 Then
                        SetFocusAPI IPAddressEditHandle(2)
                        SendMessage IPAddressEditHandle(2), EM_SETSEL, 0, ByVal -1&
                    End If
                Case IpaAutoSelectThird
                    If IPAddressEditHandle(3) <> 0 Then
                        SetFocusAPI IPAddressEditHandle(3)
                        SendMessage IPAddressEditHandle(3), EM_SETSEL, 0, ByVal -1&
                    End If
                Case IpaAutoSelectFourth
                    If IPAddressEditHandle(4) <> 0 Then
                        SetFocusAPI IPAddressEditHandle(4)
                        SendMessage IPAddressEditHandle(4), EM_SETSEL, 0, ByVal -1&
                    End If
                Case IpaAutoSelectBlank
                    Dim BlankItem As Integer
                    BlankItem = GetBlankItem()
                    If IPAddressEditHandle(BlankItem) <> 0 Then
                        SetFocusAPI IPAddressEditHandle(BlankItem)
                        SendMessage IPAddressEditHandle(BlankItem), EM_SETSEL, 0, ByVal -1&
                    End If
            End Select
        End If
    Case WM_MOUSELEAVE, WM_NCMOUSEMOVE
        If wMsg = WM_NCMOUSEMOVE And IPAddressMouseOver(1) = False Then Exit Function
        Dim TME As TRACKMOUSEEVENTSTRUCT
        With TME
        .cbSize = LenB(TME)
        .hWndTrack = hWnd
        .dwFlags = TME_LEAVE Or TME_NONCLIENT
        End With
        TrackMouseEvent TME
    Case WM_NCMOUSELEAVE
        IPAddressMouseOver(0) = False
        If IPAddressMouseOver(1) = True Then
            Dim Pos As Long, hWndFromPoint As Long
            Pos = GetMessagePos()
            hWndFromPoint = WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos))
            If (hWndFromPoint <> IPAddressEditHandle(1) Or IPAddressEditHandle(1) = 0) And (hWndFromPoint <> IPAddressEditHandle(2) Or IPAddressEditHandle(2) = 0) And (hWndFromPoint <> IPAddressEditHandle(3) Or IPAddressEditHandle(3) = 0) And (hWndFromPoint <> IPAddressEditHandle(4) Or IPAddressEditHandle(4) = 0) Then
                IPAddressMouseOver(1) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function

Private Function WindowProcEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Dim SelStart As Long, SelEnd As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd And (wParam <> IPAddressEditHandle(1) Or IPAddressEditHandle(1) = 0) And (wParam <> IPAddressEditHandle(2) Or IPAddressEditHandle(2) = 0) And (wParam <> IPAddressEditHandle(3) Or IPAddressEditHandle(3) = 0) And (wParam <> IPAddressEditHandle(4) Or IPAddressEditHandle(4) = 0) Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
        CheckMinMaxFromWindow hWnd
    Case WM_LBUTTONDOWN
        If IPAddressEditHandle(1) = 0 Or IPAddressEditHandle(2) = 0 Or IPAddressEditHandle(3) = 0 Or IPAddressEditHandle(4) = 0 Then
            If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
        Else
            Select Case GetFocus()
                Case IPAddressEditHandle(1), IPAddressEditHandle(2), IPAddressEditHandle(3), IPAddressEditHandle(4)
                Case Else
                    UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
            End Select
        End If
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
                Select Case KeyCode
                    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyBack
                        Dim Shift As Integer
                        Shift = GetShiftStateFromMsg()
                        Select Case KeyCode
                            Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                                If (Shift And vbCtrlMask) = 0 Then SendMessage hWnd, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                            Case vbKeyBack
                                SendMessage hWnd, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                        End Select
                        If SelStart = SelEnd Then
                            Dim Item As Integer
                            Item = dwRefData
                            Select Case KeyCode
                                Case vbKeyUp, vbKeyLeft
                                    If Item > 1 Then
                                        If (Shift And vbCtrlMask) <> 0 Then
                                            Item = Item - 1
                                        Else
                                            If IPAddressRTLReading(Item) = False Then
                                                If SelEnd = 0 Then Item = Item - 1
                                            Else
                                                If SelEnd = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&) Then Item = Item - 1
                                            End If
                                        End If
                                    End If
                                Case vbKeyDown, vbKeyRight
                                    If Item < 4 Then
                                        If (Shift And vbCtrlMask) <> 0 Then
                                            Item = Item + 1
                                        Else
                                            If IPAddressRTLReading(Item) = False Then
                                                If SelEnd = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&) Then Item = Item + 1
                                            Else
                                                If SelEnd = 0 Then Item = Item + 1
                                            End If
                                        End If
                                    End If
                                Case vbKeyHome
                                    Item = 1
                                Case vbKeyEnd
                                    Item = 4
                                Case vbKeyBack
                                    If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) = 0 Then
                                        If Item > 1 Then
                                            If SelEnd = 0 Then Item = Item - 1
                                        End If
                                    End If
                            End Select
                            If Item <> dwRefData Then
                                If CheckMinMaxFromWindow(IPAddressEditHandle(dwRefData)) = False Then
                                    If GetFocus() <> IPAddressEditHandle(Item) And IPAddressEditHandle(Item) <> 0 Then
                                        UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
                                        SetFocusAPI IPAddressEditHandle(Item)
                                        Select Case KeyCode
                                            Case vbKeyUp, vbKeyLeft
                                                If (Shift And vbCtrlMask) <> 0 Then
                                                    SendMessage IPAddressEditHandle(Item), EM_SETSEL, 0, ByVal -1&
                                                Else
                                                    If IPAddressRTLReading(Item) = False Then
                                                        SelEnd = SendMessage(IPAddressEditHandle(Item), WM_GETTEXTLENGTH, 0, ByVal 0&)
                                                        SendMessage IPAddressEditHandle(Item), EM_SETSEL, SelEnd, ByVal SelEnd
                                                    Else
                                                        SendMessage IPAddressEditHandle(Item), EM_SETSEL, 0, ByVal 0&
                                                    End If
                                                End If
                                            Case vbKeyDown, vbKeyRight
                                                If (Shift And vbCtrlMask) <> 0 Then
                                                    SendMessage IPAddressEditHandle(Item), EM_SETSEL, 0, ByVal -1&
                                                Else
                                                    If IPAddressRTLReading(Item) = False Then
                                                        SendMessage IPAddressEditHandle(Item), EM_SETSEL, 0, ByVal 0&
                                                    Else
                                                        SelEnd = SendMessage(IPAddressEditHandle(Item), WM_GETTEXTLENGTH, 0, ByVal 0&)
                                                        SendMessage IPAddressEditHandle(Item), EM_SETSEL, SelEnd, ByVal SelEnd
                                                    End If
                                                End If
                                            Case vbKeyHome
                                                SendMessage IPAddressEditHandle(Item), EM_SETSEL, 0, ByVal 0&
                                            Case vbKeyEnd
                                                SelEnd = SendMessage(IPAddressEditHandle(Item), WM_GETTEXTLENGTH, 0, ByVal 0&)
                                                SendMessage IPAddressEditHandle(Item), EM_SETSEL, SelEnd, ByVal SelEnd
                                            Case vbKeyBack
                                                SelEnd = SendMessage(IPAddressEditHandle(Item), WM_GETTEXTLENGTH, 0, ByVal 0&)
                                                SendMessage IPAddressEditHandle(Item), EM_SETSEL, SelEnd, ByVal SelEnd
                                                If SelEnd > 0 Then PostMessage IPAddressEditHandle(Item), WM_KEYDOWN, vbKeyBack, ByVal 0&
                                        End Select
                                    End If
                                End If
                            End If
                        End If
                End Select
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            IPAddressCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If IPAddressCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(IPAddressCharCodeCache And &HFFFF&)
            IPAddressCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        If (wParam And &HFFFF&) <> 0 And KeyChar = 0 Then
            Exit Function
        Else
            wParam = CIntToUInt(KeyChar)
        End If
        Select Case wParam
            Case 32, 46 ' " ", "."
                SendMessage hWnd, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                If SelStart = SelEnd And SelStart > 0 Then
                    If dwRefData < 4 Then
                        ' CheckMinMaxFromWindow validation not necessary as no modification happens.
                        ' Just change focus to the next edit control.
                        If GetFocus() <> IPAddressEditHandle(dwRefData + 1) And IPAddressEditHandle(dwRefData + 1) <> 0 Then
                            UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
                            SetFocusAPI IPAddressEditHandle(dwRefData + 1)
                            SendMessage IPAddressEditHandle(dwRefData + 1), EM_SETSEL, 0, ByVal -1&
                        End If
                    Else
                        wParam = 0 ' Beep
                    End If
                End If
                If wParam <> 0 Then Exit Function ' Avoid ES_NUMBER style balloon tip
        End Select
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcEdit = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(wParam)
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcEdit = 0
        End If
        Exit Function
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_CONTEXTMENU
        If wParam = hWnd Then
            Dim P1 As POINTAPI, Handled As Boolean
            P1.X = Get_X_lParam(lParam)
            P1.Y = Get_Y_lParam(lParam)
            If P1.X = -1 And P1.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(Handled, -1, -1)
            Else
                ScreenToClient UserControl.hWnd, P1
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P1.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P1.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
    Case WM_PASTE
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
    Case WM_STYLECHANGING, WM_STYLECHANGED
        Dim dwStyleOld As Long, dwStyleNew As Long
        If wMsg = WM_STYLECHANGING And wParam = GWL_STYLE Then
            CopyMemory dwStyleNew, ByVal UnsignedAdd(lParam, 4), 4
            dwStyleOld = dwStyleNew
            If (dwStyleNew And ES_LEFT) = ES_LEFT Then dwStyleNew = dwStyleNew And Not ES_LEFT
            If (dwStyleNew And ES_CENTER) = ES_CENTER Then dwStyleNew = dwStyleNew And Not ES_CENTER
            If (dwStyleNew And ES_RIGHT) = ES_RIGHT Then dwStyleNew = dwStyleNew And Not ES_RIGHT
            ' Enforcing ES_CENTER style and circumvent unwanted modification.
            ' For example, when changing the right-to-left reading in the default context menu.
            dwStyleNew = dwStyleNew Or ES_CENTER
            If dwStyleOld <> dwStyleNew Then CopyMemory ByVal UnsignedAdd(lParam, 4), dwStyleNew, 4
        ElseIf wMsg = WM_STYLECHANGED And wParam = GWL_EXSTYLE Then
            CopyMemory dwStyleNew, ByVal UnsignedAdd(lParam, 4), 4
            IPAddressRTLReading(dwRefData) = CBool((dwStyleNew And WS_EX_RTLREADING) = WS_EX_RTLREADING)
        End If
End Select
WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_SETFOCUS
        IPAddressEditFocusHwnd = hWnd
        
        #If ImplementThemedBorder = True Then
        
        If PropBorderStyle = CCBorderStyleSunken And PropVisualStyles = True Then
            If IPAddressEnabledVisualStyles = True Then SetWindowPos UserControl.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
        End If
        
        #End If
        
        If dwRefData <> IPAddressSelectedItem Then
            IPAddressSelectedItem = dwRefData
            RaiseEvent SelChange
        End If
    Case WM_KILLFOCUS
        IPAddressEditFocusHwnd = 0
        
        #If ImplementThemedBorder = True Then
        
        If PropBorderStyle = CCBorderStyleSunken And PropVisualStyles = True Then
            If wParam <> UserControl.hWnd Then ' Avoid flicker
                If IPAddressEnabledVisualStyles = True Then SetWindowPos UserControl.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
            End If
        End If
        
        #End If
        
    Case WM_CHAR
        Select Case wParam
            Case 48 To 57 ' "0-9"
                SendMessage hWnd, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                If SelStart = 3 And SelEnd = 3 Then
                    If CheckMinMaxFromWindow(hWnd) = False Then
                        If dwRefData < 4 Then
                            If GetFocus() <> IPAddressEditHandle(dwRefData + 1) And IPAddressEditHandle(dwRefData + 1) <> 0 Then
                                UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
                                SetFocusAPI IPAddressEditHandle(dwRefData + 1)
                                SendMessage IPAddressEditHandle(dwRefData + 1), EM_SETSEL, 0, ByVal -1&
                            End If
                        End If
                    End If
                End If
        End Select
    Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        RaiseEvent DblClick
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim P2 As POINTAPI
        P2.X = Get_X_lParam(lParam)
        P2.Y = Get_Y_lParam(lParam)
        MapWindowPoints hWnd, UserControl.hWnd, P2, 1
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(P2.X, vbPixels, vbTwips)
        Y = UserControl.ScaleY(P2.Y, vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                IPAddressIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                IPAddressIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                IPAddressIsClick = True
            Case WM_MOUSEMOVE
                If (IPAddressEditMouseOver(dwRefData) = False And PropMouseTrack = True) Or (IPAddressMouseOver(1) = False And PropMouseTrack = True) Then
                    If IPAddressEditMouseOver(dwRefData) = False And PropMouseTrack = True Then IPAddressEditMouseOver(dwRefData) = True
                    If IPAddressMouseOver(1) = False And PropMouseTrack = True Then
                        IPAddressMouseOver(1) = True
                        RaiseEvent MouseEnter
                    End If
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
                If IPAddressIsClick = True Then
                    IPAddressIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        IPAddressEditMouseOver(dwRefData) = False
        If IPAddressMouseOver(1) = True Then
            Dim Pos As Long, hWndFromPoint As Long
            Pos = GetMessagePos()
            hWndFromPoint = WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos))
            If hWndFromPoint <> UserControl.hWnd And (hWndFromPoint <> IPAddressEditHandle(1) Or IPAddressEditHandle(1) = 0) And (hWndFromPoint <> IPAddressEditHandle(2) Or IPAddressEditHandle(2) = 0) And (hWndFromPoint <> IPAddressEditHandle(3) Or IPAddressEditHandle(3) = 0) And (hWndFromPoint <> IPAddressEditHandle(4) Or IPAddressEditHandle(4) = 0) Then
                IPAddressMouseOver(1) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function
