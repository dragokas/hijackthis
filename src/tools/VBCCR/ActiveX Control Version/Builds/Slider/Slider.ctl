VERSION 5.00
Begin VB.UserControl Slider 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "Slider.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "Slider.ctx":004D
End
Attribute VB_Name = "Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private SldOrientationHorizontal, SldOrientationVertical
Private SldTipSideAboveLeft, SldTipSideBelowRight
Private SldTickStyleBottomRight, SldTickStyleTopLeft, SldTickStyleBoth, SldTickStyleNone
Private SldDrawModeNormal, SldDrawModeOwnerDraw
Private SldOwnerDrawItemTics, SldOwnerDrawItemThumb, SldOwnerDrawItemChannel
#End If
Public Enum SldOrientationConstants
SldOrientationHorizontal = 0
SldOrientationVertical = 1
End Enum
Public Enum SldTipSideConstants
SldTipSideAboveLeft = 0
SldTipSideBelowRight = 1
End Enum
Public Enum SldTickStyleConstants
SldTickStyleBottomRight = 0
SldTickStyleTopLeft = 1
SldTickStyleBoth = 2
SldTickStyleNone = 3
End Enum
Public Enum SldDrawModeConstants
SldDrawModeNormal = 0
SldDrawModeOwnerDraw = 1
End Enum
Private Const TBCD_TICS As Long = &H1
Private Const TBCD_THUMB As Long = &H2
Private Const TBCD_CHANNEL As Long = &H3
Public Enum SldOwnerDrawItemConstants
SldOwnerDrawItemTics = TBCD_TICS
SldOwnerDrawItemThumb = TBCD_THUMB
SldOwnerDrawItemChannel = TBCD_CHANNEL
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Const CDDS_PREPAINT As Long = &H1
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM + 1)
Private Const CDRF_DODEFAULT As Long = &H0
Private Const CDRF_SKIPDEFAULT As Long = &H4
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20
Private Type NMCUSTOMDRAW
hdr As NMHDR
dwDrawStage As Long
hDC As Long
RC As RECT
dwItemSpec As Long
uItemState As Long
lItemlParam As Long
End Type
Private Type NMTTDISPINFO
hdr As NMHDR
lpszText As Long
szText(0 To ((80 * 2) - 1)) As Byte
hInst As Long
uFlags As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event Scroll()
Attribute Scroll.VB_Description = "Occurs when repositioning."
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event ContextMenu(ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event ModifyTipText(ByRef Text As String)
Attribute ModifyTipText.VB_Description = "Occurs if the slider control is about to display a position tip. This is a request to modify the text to display. This will only occur if the show tips property is set to true."
Public Event ItemDraw(ByVal Item As SldOwnerDrawItemConstants, ByRef Cancel As Boolean, ByVal ItemState As Long, ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Attribute ItemDraw.VB_Description = "Occurs when a visual aspect of an owner-drawn slider has changed."
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
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const ICC_BAR_CLASSES As Long = &H20
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const HWND_DESKTOP As Long = &H0
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const SW_HIDE As Long = &H0
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114
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
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_PAINT As Long = &HF
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const TB_THUMBPOSITION As Long = 4
Private Const TB_THUMBTRACK As Long = 5
Private Const TB_ENDTRACK As Long = 8
Private Const WM_USER As Long = &H400
Private Const TBM_GETPOS As Long = (WM_USER)
Private Const TBM_GETRANGEMIN As Long = (WM_USER + 1)
Private Const TBM_GETRANGEMAX As Long = (WM_USER + 2)
Private Const TBM_GETTIC As Long = (WM_USER + 3)
Private Const TBM_SETTIC As Long = (WM_USER + 4)
Private Const TBM_SETPOS As Long = (WM_USER + 5)
Private Const TBM_SETRANGE As Long = (WM_USER + 6) ' 16 bit
Private Const TBM_SETRANGEMIN As Long = (WM_USER + 7)
Private Const TBM_SETRANGEMAX As Long = (WM_USER + 8)
Private Const TBM_CLEARTICS As Long = (WM_USER + 9)
Private Const TBM_SETSEL As Long = (WM_USER + 10)
Private Const TBM_SETSELSTART As Long = (WM_USER + 11)
Private Const TBM_SETSELEND As Long = (WM_USER + 12)
Private Const TBM_GETPTICS As Long = (WM_USER + 14)
Private Const TBM_GETTICPOS As Long = (WM_USER + 15)
Private Const TBM_GETNUMTICS As Long = (WM_USER + 16)
Private Const TBM_GETSELSTART As Long = (WM_USER + 17)
Private Const TBM_GETSELEND As Long = (WM_USER + 18)
Private Const TBM_CLEARSEL As Long = (WM_USER + 19)
Private Const TBM_SETTICFREQ As Long = (WM_USER + 20)
Private Const TBM_SETPAGESIZE As Long = (WM_USER + 21)
Private Const TBM_GETPAGESIZE As Long = (WM_USER + 22)
Private Const TBM_SETLINESIZE As Long = (WM_USER + 23)
Private Const TBM_GETLINESIZE As Long = (WM_USER + 24)
Private Const TBM_GETTHUMBRECT As Long = (WM_USER + 25)
Private Const TBM_GETCHANNELRECT As Long = (WM_USER + 26)
Private Const TBM_SETTHUMBLENGTH As Long = (WM_USER + 27)
Private Const TBM_GETTHUMBLENGTH As Long = (WM_USER + 28)
Private Const TBM_SETTOOLTIPS As Long = (WM_USER + 29)
Private Const TBM_GETTOOLTIPS As Long = (WM_USER + 30)
Private Const TBM_SETTIPSIDE As Long = (WM_USER + 31)
Private Const TBM_SETBUDDY As Long = (WM_USER + 32)
Private Const TBM_GETBUDDY As Long = (WM_USER + 33)
Private Const TBS_AUTOTICKS As Long = &H1
Private Const TBS_VERT As Long = &H2
Private Const TBS_HORZ As Long = &H0
Private Const TBS_TOP As Long = &H4
Private Const TBS_BOTTOM As Long = &H0
Private Const TBS_LEFT As Long = &H4
Private Const TBS_RIGHT As Long = &H0
Private Const TBS_BOTH As Long = &H8
Private Const TBS_NOTICKS As Long = &H10
Private Const TBS_ENABLESELRANGE As Long = &H20
Private Const TBS_FIXEDLENGTH As Long = &H40
Private Const TBS_NOTHUMB As Long = &H80
Private Const TBS_TOOLTIPS As Long = &H100
Private Const TBS_REVERSED As Long = &H200
Private Const TBS_DOWNISLEFT As Long = &H400
Private Const TBTS_TOP As Long = 0
Private Const TBTS_LEFT As Long = 1
Private Const TBTS_BOTTOM As Long = 2
Private Const TBTS_RIGHT As Long = 3
Private Const NM_FIRST As Long = 0
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Private Const TTF_RTLREADING As Long = &H4
Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFOA As Long = (TTN_FIRST - 0)
Private Const TTN_GETDISPINFOW As Long = (TTN_FIRST - 10)
Private Const TTN_GETDISPINFO As Long = TTN_GETDISPINFOW
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private SliderHandle As Long, SliderToolTipHandle As Long
Private SliderTransparentBrush As Long
Private SliderCharCodeCache As Long
Private SliderIsClick As Boolean
Private SliderMouseOver As Boolean
Private SliderDesignMode As Boolean
Private SliderMaxExtentX As Long
Private SliderMaxExtentY As Long
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropMin As Long, PropMax As Long
Private PropValue As Long
Private PropTickFrequency As Long
Private PropOrientation As SldOrientationConstants
Private PropSmallChange As Long, PropLargeChange As Long
Private PropTickStyle As SldTickStyleConstants
Private PropShowTip As Boolean
Private PropTipSide As SldTipSideConstants
Private PropSelectRange As Boolean
Private PropSelStart As Long, PropSelLength As Long
Private PropTransparent As Boolean
Private PropHideThumb As Boolean
Private PropReversed As Boolean
Private PropDrawMode As SldDrawModeConstants

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
Call ComCtlsInitCC(ICC_BAR_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
SliderMaxExtentX = 45 * PixelsPerDIP_X()
SliderMaxExtentY = 45 * PixelsPerDIP_Y()
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
SliderDesignMode = Not Ambient.UserMode
On Error GoTo 0
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropMin = 0
PropMax = 10
PropValue = 0
PropTickFrequency = 1
PropOrientation = SldOrientationHorizontal
PropSmallChange = 1
PropLargeChange = 2
PropShowTip = True
PropTipSide = SldTipSideAboveLeft
PropTickStyle = SldTickStyleBottomRight
PropSelectRange = False
PropSelStart = 0
PropSelLength = 0
PropTransparent = False
PropHideThumb = False
PropReversed = False
PropDrawMode = SldDrawModeNormal
Call CreateSlider
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
SliderDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropMin = .ReadProperty("Min", 0)
PropMax = .ReadProperty("Max", 10)
PropValue = .ReadProperty("Value", 0)
PropTickFrequency = .ReadProperty("TickFrequency", 1)
PropOrientation = .ReadProperty("Orientation", SldOrientationHorizontal)
PropSmallChange = .ReadProperty("SmallChange", 1)
PropLargeChange = .ReadProperty("LargeChange", 2)
PropTickStyle = .ReadProperty("TickStyle", SldTickStyleBottomRight)
PropShowTip = .ReadProperty("ShowTip", True)
PropTipSide = .ReadProperty("TipSide", SldTipSideAboveLeft)
PropSelectRange = .ReadProperty("SelectRange", False)
PropSelStart = .ReadProperty("SelStart", 0)
PropSelLength = .ReadProperty("SelLength", 0)
PropTransparent = .ReadProperty("Transparent", False)
PropHideThumb = .ReadProperty("HideThumb", False)
PropReversed = .ReadProperty("Reversed", False)
PropDrawMode = .ReadProperty("DrawMode", SldDrawModeNormal)
End With
Call CreateSlider
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "Min", PropMin, 0
.WriteProperty "Max", PropMax, 10
.WriteProperty "Value", PropValue, 0
.WriteProperty "TickFrequency", PropTickFrequency, 1
.WriteProperty "Orientation", PropOrientation, SldOrientationHorizontal
.WriteProperty "SmallChange", PropSmallChange, 1
.WriteProperty "LargeChange", PropLargeChange, 2
.WriteProperty "TickStyle", PropTickStyle, SldTickStyleBottomRight
.WriteProperty "ShowTip", PropShowTip, True
.WriteProperty "TipSide", PropTipSide, SldTipSideAboveLeft
.WriteProperty "SelectRange", PropSelectRange, False
.WriteProperty "SelStart", PropSelStart, 0
.WriteProperty "SelLength", PropSelLength, 0
.WriteProperty "Transparent", PropTransparent, False
.WriteProperty "HideThumb", PropHideThumb, False
.WriteProperty "Reversed", PropReversed, False
.WriteProperty "DrawMode", PropDrawMode, SldDrawModeNormal
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
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
Dim Width As Long, Height As Long
Width = .ScaleWidth
Height = .ScaleHeight
Select Case PropOrientation
    Case SldOrientationHorizontal
        If Height > SliderMaxExtentY Then Height = SliderMaxExtentY
    Case SldOrientationVertical
        If Width > SliderMaxExtentX Then Width = SliderMaxExtentX
End Select
If SliderHandle <> 0 Then
    If PropTransparent = True Then
        MoveWindow SliderHandle, 0, 0, Width, Height, 0
        If SliderTransparentBrush <> 0 Then
            DeleteObject SliderTransparentBrush
            SliderTransparentBrush = 0
        End If
        RedrawWindow SliderHandle, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
    Else
        MoveWindow SliderHandle, 0, 0, Width, Height, 1
    End If
End If
.Extender.Move .Extender.Left, .Extender.Top, .ScaleX(Width, vbPixels, vbContainerSize), .ScaleY(Height, vbPixels, vbContainerSize)
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroySlider
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
hWnd = SliderHandle
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
If SliderHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles SliderHandle
    Else
        RemoveVisualStyles SliderHandle
    End If
    Call SetVisualStylesToolTip
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
If SliderHandle <> 0 Then Call ReCreateSlider
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If SliderHandle <> 0 Then EnableWindow SliderHandle, IIf(Value = True, 1, 0)
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
If SliderDesignMode = False Then Call RefreshMousePointer
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
        If SliderDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If SliderDesignMode = False Then Call RefreshMousePointer
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
If SliderDesignMode = False Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
    dwMask = 0
End If
If SliderHandle <> 0 Then Call ReCreateSlider
UserControl.PropertyChanged "RightToLeft"
End Property

Public Property Get RightToLeftLayout() As Boolean
Attribute RightToLeftLayout.VB_Description = "Returns/sets a value indicating if right-to-left mirror placement is turned on."
RightToLeftLayout = PropRightToLeftLayout
End Property

Public Property Let RightToLeftLayout(ByVal Value As Boolean)
PropRightToLeftLayout = Value
Me.RightToLeft = PropRightToLeft
UserControl.PropertyChanged "RightToLeftLayout"
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

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum value."
If SliderHandle <> 0 Then
    Min = SendMessage(SliderHandle, TBM_GETRANGEMIN, 0, ByVal 0&)
Else
    Min = PropMin
End If
End Property

Public Property Let Min(ByVal Value As Long)
If Value < Me.Max Then
    PropMin = Value
    If PropValue < PropMin Then PropValue = PropMin
    If PropMin > PropSelStart Then
        PropSelStart = PropMin
        PropSelLength = 0
    End If
Else
    If SliderDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If SliderHandle <> 0 Then
    If PropSelectRange = True Then
        SendMessage SliderHandle, TBM_SETRANGEMIN, 0, ByVal PropMin
        SendMessage SliderHandle, TBM_SETSELSTART, 0, ByVal PropSelStart
        SendMessage SliderHandle, TBM_SETSELEND, 1, ByVal (PropSelStart + PropSelLength)
    Else
        SendMessage SliderHandle, TBM_SETRANGEMIN, 1, ByVal PropMin
    End If
End If
UserControl.PropertyChanged "Min"
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum value."
If SliderHandle <> 0 Then
    Max = SendMessage(SliderHandle, TBM_GETRANGEMAX, 0, ByVal 0&)
Else
    Max = PropMax
End If
End Property

Public Property Let Max(ByVal Value As Long)
If Value > Me.Min Then
    PropMax = Value
    If PropValue > PropMax Then PropValue = PropMax
    If PropMax < PropSelStart Then
        PropSelStart = PropMax
        PropSelLength = 0
    End If
Else
    If SliderDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If SliderHandle <> 0 Then
    If PropSelectRange = True Then
        SendMessage SliderHandle, TBM_SETRANGEMAX, 0, ByVal PropMax
        SendMessage SliderHandle, TBM_SETSELSTART, 0, ByVal PropSelStart
        SendMessage SliderHandle, TBM_SETSELEND, 1, ByVal (PropSelStart + PropSelLength)
    Else
        SendMessage SliderHandle, TBM_SETRANGEMAX, 1, ByVal PropMax
    End If
End If
UserControl.PropertyChanged "Max"
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the current position."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "123c"
If SliderHandle <> 0 Then
    Value = SendMessage(SliderHandle, TBM_GETPOS, 0, ByVal 0&)
Else
    Value = PropValue
End If
End Property

Public Property Let Value(ByVal NewValue As Long)
If NewValue > Me.Max Then
    NewValue = Me.Max
ElseIf NewValue < Me.Min Then
    NewValue = Me.Min
End If
Dim Changed As Boolean
Changed = CBool(Me.Value <> NewValue)
PropValue = NewValue
If SliderHandle <> 0 Then SendMessage SliderHandle, TBM_SETPOS, 1, ByVal PropValue
UserControl.PropertyChanged "Value"
If Changed = True Then
    On Error Resume Next
    UserControl.Extender.DataChanged = True
    On Error GoTo 0
    RaiseEvent Change
End If
End Property

Public Property Get TickFrequency() As Long
Attribute TickFrequency.VB_Description = "Returns/sets the ratio of ticks; 1tick every n increments."
TickFrequency = PropTickFrequency
End Property

Public Property Let TickFrequency(ByVal Value As Long)
If Value > 0 Then
    PropTickFrequency = Value
Else
    If SliderDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If SliderHandle <> 0 Then SendMessage SliderHandle, TBM_SETTICFREQ, PropTickFrequency, ByVal 0&
UserControl.PropertyChanged "TickFrequency"
End Property

Public Property Get Orientation() As SldOrientationConstants
Attribute Orientation.VB_Description = "Returns/sets the orientation."
Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As SldOrientationConstants)
Dim Swap(0 To 1) As Long
Select Case Value
    Case SldOrientationHorizontal, SldOrientationVertical
        If PropOrientation <> Value Then
            Swap(0) = UserControl.ScaleHeight
            Swap(1) = UserControl.ScaleWidth
        Else
            Swap(0) = -1
            Swap(1) = -1
        End If
        PropOrientation = Value
    Case Else
        Err.Raise 380
End Select
If SliderHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(SliderHandle, GWL_STYLE)
    If (dwStyle And TBS_HORZ) = TBS_HORZ Then dwStyle = dwStyle And Not TBS_HORZ
    If (dwStyle And TBS_VERT) = TBS_VERT Then dwStyle = dwStyle And Not TBS_VERT
    If (dwStyle And TBS_BOTTOM) = TBS_BOTTOM Then dwStyle = dwStyle And Not TBS_BOTTOM
    If (dwStyle And TBS_RIGHT) = TBS_RIGHT Then dwStyle = dwStyle And Not TBS_RIGHT
    If (dwStyle And TBS_TOP) = TBS_TOP Then dwStyle = dwStyle And Not TBS_TOP
    If (dwStyle And TBS_LEFT) = TBS_LEFT Then dwStyle = dwStyle And Not TBS_LEFT
    If (dwStyle And TBS_BOTH) = TBS_BOTH Then dwStyle = dwStyle And Not TBS_BOTH
    If (dwStyle And TBS_NOTICKS) = TBS_NOTICKS Then dwStyle = dwStyle And Not TBS_NOTICKS
    If PropOrientation = SldOrientationHorizontal Then
        dwStyle = dwStyle Or TBS_HORZ
    ElseIf PropOrientation = SldOrientationVertical Then
        dwStyle = dwStyle Or TBS_VERT
    End If
    Select Case PropTickStyle
        Case SldTickStyleBottomRight
            If PropOrientation = SldOrientationHorizontal Then
                dwStyle = dwStyle Or TBS_BOTTOM
            ElseIf PropOrientation = SldOrientationVertical Then
                dwStyle = dwStyle Or TBS_RIGHT
            End If
        Case SldTickStyleTopLeft
            If PropOrientation = SldOrientationHorizontal Then
                dwStyle = dwStyle Or TBS_TOP
            ElseIf PropOrientation = SldOrientationVertical Then
                dwStyle = dwStyle Or TBS_LEFT
            End If
        Case SldTickStyleBoth
            dwStyle = dwStyle Or TBS_BOTH
        Case SldTickStyleNone
            dwStyle = dwStyle Or TBS_NOTICKS
    End Select
    SetWindowLong SliderHandle, GWL_STYLE, dwStyle
    If Swap(0) > -1 And Swap(1) > -1 Then
        With UserControl
        .Extender.Move .Extender.Left, .Extender.Top, .ScaleX(Swap(0), vbPixels, vbContainerSize), .ScaleY(Swap(1), vbPixels, vbContainerSize)
        End With
    End If
End If
UserControl.PropertyChanged "Orientation"
End Property

Public Property Get SmallChange() As Long
Attribute SmallChange.VB_Description = "Returns/sets the number of logical position moves in response to keyboard input from the arrow keys."
If SliderHandle <> 0 Then
    SmallChange = SendMessage(SliderHandle, TBM_GETLINESIZE, 0, ByVal 0&)
Else
    SmallChange = PropSmallChange
End If
End Property

Public Property Let SmallChange(ByVal Value As Long)
PropSmallChange = Value
If SliderHandle <> 0 Then SendMessage SliderHandle, TBM_SETLINESIZE, 0, ByVal PropSmallChange
UserControl.PropertyChanged "SmallChange"
End Property

Public Property Get LargeChange() As Long
Attribute LargeChange.VB_Description = "Returns/sets the number of logical position moves in response to keyboard input from the page up or page down keys."
If SliderHandle <> 0 Then
    LargeChange = SendMessage(SliderHandle, TBM_GETPAGESIZE, 0, ByVal 0&)
Else
    LargeChange = PropLargeChange
End If
End Property

Public Property Let LargeChange(ByVal Value As Long)
PropLargeChange = Value
If SliderHandle <> 0 Then SendMessage SliderHandle, TBM_SETPAGESIZE, 0, ByVal PropLargeChange
UserControl.PropertyChanged "LargeChange"
End Property

Public Property Get TickStyle() As SldTickStyleConstants
Attribute TickStyle.VB_Description = "Returns/sets the style (or positioning) of the tick marks displayed."
TickStyle = PropTickStyle
End Property

Public Property Let TickStyle(ByVal Value As SldTickStyleConstants)
Select Case Value
    Case SldTickStyleBottomRight, SldTickStyleTopLeft, SldTickStyleBoth, SldTickStyleNone
        PropTickStyle = Value
    Case Else
        Err.Raise 380
End Select
If SliderHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(SliderHandle, GWL_STYLE)
    If (dwStyle And TBS_BOTTOM) = TBS_BOTTOM Then dwStyle = dwStyle And Not TBS_BOTTOM
    If (dwStyle And TBS_RIGHT) = TBS_RIGHT Then dwStyle = dwStyle And Not TBS_RIGHT
    If (dwStyle And TBS_TOP) = TBS_TOP Then dwStyle = dwStyle And Not TBS_TOP
    If (dwStyle And TBS_LEFT) = TBS_LEFT Then dwStyle = dwStyle And Not TBS_LEFT
    If (dwStyle And TBS_BOTH) = TBS_BOTH Then dwStyle = dwStyle And Not TBS_BOTH
    If (dwStyle And TBS_NOTICKS) = TBS_NOTICKS Then dwStyle = dwStyle And Not TBS_NOTICKS
    Select Case PropTickStyle
        Case SldTickStyleBottomRight
            If PropOrientation = SldOrientationHorizontal Then
                dwStyle = dwStyle Or TBS_BOTTOM
            ElseIf PropOrientation = SldOrientationVertical Then
                dwStyle = dwStyle Or TBS_RIGHT
            End If
        Case SldTickStyleTopLeft
            If PropOrientation = SldOrientationHorizontal Then
                dwStyle = dwStyle Or TBS_TOP
            ElseIf PropOrientation = SldOrientationVertical Then
                dwStyle = dwStyle Or TBS_LEFT
            End If
        Case SldTickStyleBoth
            dwStyle = dwStyle Or TBS_BOTH
        Case SldTickStyleNone
            dwStyle = dwStyle Or TBS_NOTICKS
    End Select
    SetWindowLong SliderHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "TickStyle"
End Property

Public Property Get ShowTip() As Boolean
Attribute ShowTip.VB_Description = "Returns/sets a value that determines whether a position tip will be displayed or not."
ShowTip = PropShowTip
End Property

Public Property Let ShowTip(ByVal Value As Boolean)
PropShowTip = Value
If SliderHandle <> 0 And SliderDesignMode = False Then
    If PropShowTip = False Then
        SendMessage SliderHandle, TBM_SETTOOLTIPS, 0, ByVal 0&
    Else
        If SliderToolTipHandle = 0 Then Call ReCreateSlider
        If SliderToolTipHandle <> 0 Then SendMessage SliderHandle, TBM_SETTOOLTIPS, SliderToolTipHandle, ByVal 0&
    End If
End If
UserControl.PropertyChanged "ShowTip"
End Property

Public Property Get TipSide() As SldTipSideConstants
Attribute TipSide.VB_Description = "Returns/sets a value representing the location at which to display the position tip. Only applicable if the show tip property is set to true."
TipSide = PropTipSide
End Property

Public Property Let TipSide(ByVal Value As SldTipSideConstants)
Select Case Value
    Case SldTipSideAboveLeft, SldTipSideBelowRight
        PropTipSide = Value
    Case Else
        Err.Raise 380
End Select
If SliderHandle <> 0 Then
    Dim SetVal As Long
    If PropOrientation = SldOrientationHorizontal Then
        If PropTipSide = SldTipSideAboveLeft Then
            SetVal = TBTS_TOP
        ElseIf PropTipSide = SldTipSideBelowRight Then
            SetVal = TBTS_BOTTOM
        End If
    ElseIf PropOrientation = SldOrientationVertical Then
        If PropTipSide = SldTipSideAboveLeft Then
            SetVal = TBTS_LEFT
        ElseIf PropTipSide = SldTipSideBelowRight Then
            SetVal = TBTS_RIGHT
        End If
    End If
    SendMessage SliderHandle, TBM_SETTIPSIDE, SetVal, ByVal 0&
End If
UserControl.PropertyChanged "TipSide"
End Property

Public Property Get SelectRange() As Boolean
Attribute SelectRange.VB_Description = "Returns/sets whether or not a trackbar can have a select range."
SelectRange = PropSelectRange
End Property

Public Property Let SelectRange(ByVal Value As Boolean)
If Not PropSelectRange = Value Then
    PropSelStart = Me.Value
    PropSelLength = 0
End If
PropSelectRange = Value
If SliderHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(SliderHandle, GWL_STYLE)
    If PropSelectRange = True Then
        If Not (dwStyle And TBS_ENABLESELRANGE) = TBS_ENABLESELRANGE Then
            SetWindowLong SliderHandle, GWL_STYLE, dwStyle Or TBS_ENABLESELRANGE
            SendMessage SliderHandle, TBM_SETSELSTART, 0, ByVal PropSelStart
            SendMessage SliderHandle, TBM_SETSELEND, 1, ByVal (PropSelStart + PropSelLength)
        End If
    Else
        If (dwStyle And TBS_ENABLESELRANGE) = TBS_ENABLESELRANGE Then
            SetWindowLong SliderHandle, GWL_STYLE, dwStyle And Not TBS_ENABLESELRANGE
            SendMessage SliderHandle, TBM_CLEARSEL, 1, ByVal 0&
        End If
    End If
End If
UserControl.PropertyChanged "SelectRange"
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the selection start."
If PropSelectRange = True Then
    If SliderHandle <> 0 Then
        SelStart = SendMessage(SliderHandle, TBM_GETSELSTART, 0, ByVal 0&)
    Else
        SelStart = PropSelStart
    End If
Else
    SelStart = Me.Value
End If
End Property

Public Property Let SelStart(ByVal Value As Long)
Select Case Value
    Case Me.Min To Me.Max
        PropSelStart = Value
    Case Else
        If SliderDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
If SliderHandle <> 0 Then
    If PropSelectRange = True Then
        SendMessage SliderHandle, TBM_SETSELSTART, 0, ByVal PropSelStart
        SendMessage SliderHandle, TBM_SETSELEND, 1, ByVal (PropSelStart + PropSelLength)
        PropSelLength = SendMessage(SliderHandle, TBM_GETSELEND, 0, ByVal 0&) - SendMessage(SliderHandle, TBM_GETSELSTART, 0, ByVal 0&)
    Else
        SendMessage SliderHandle, TBM_SETPOS, 1, ByVal PropSelStart
    End If
End If
UserControl.PropertyChanged "SelStart"
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the selection length."
If PropSelectRange = True Then
    If SliderHandle <> 0 Then
        SelLength = SendMessage(SliderHandle, TBM_GETSELEND, 0, ByVal 0&) - SendMessage(SliderHandle, TBM_GETSELSTART, 0, ByVal 0&)
    Else
        SelLength = PropSelLength
    End If
Else
    SelLength = 0
End If
End Property

Public Property Let SelLength(ByVal Value As Long)
If PropSelectRange = True Then
    If Value >= 0 And (PropSelStart + Value) <= Me.Max Then
        PropSelLength = Value
    Else
        If SliderDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
    If SliderHandle <> 0 Then
        SendMessage SliderHandle, TBM_SETSELSTART, 0, ByVal PropSelStart
        SendMessage SliderHandle, TBM_SETSELEND, 1, ByVal (PropSelStart + PropSelLength)
    End If
Else
    If Value <> 0 Then
        If SliderDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
UserControl.PropertyChanged "SelLength"
End Property

Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Returns/sets a value indicating if the background is a replica of the underlying background to simulate transparency. This property is ignored at design time."
Transparent = PropTransparent
End Property

Public Property Let Transparent(ByVal Value As Boolean)
PropTransparent = Value
If SliderHandle <> 0 Then Call ReCreateSlider
UserControl.PropertyChanged "Transparent"
End Property

Public Property Get HideThumb() As Boolean
Attribute HideThumb.VB_Description = "Returns/sets a value that determines whether or not the thumb marker is hidden."
HideThumb = PropHideThumb
End Property

Public Property Let HideThumb(ByVal Value As Boolean)
PropHideThumb = Value
If SliderHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(SliderHandle, GWL_STYLE)
    If PropHideThumb = True Then
        If Not (dwStyle And TBS_NOTHUMB) = TBS_NOTHUMB Then dwStyle = dwStyle Or TBS_NOTHUMB
    Else
        If (dwStyle And TBS_NOTHUMB) = TBS_NOTHUMB Then dwStyle = dwStyle And Not TBS_NOTHUMB
    End If
    SetWindowLong SliderHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "HideThumb"
End Property

Public Property Get Reversed() As Boolean
Attribute Reversed.VB_Description = "Returns/sets a value that determines whether or not to reverse the default, making down equal left and up equal right on horizontal orientation and left equal down and right equal up on vertical orientation."
Reversed = PropReversed
End Property

Public Property Let Reversed(ByVal Value As Boolean)
PropReversed = Value
If SliderHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(SliderHandle, GWL_STYLE)
    ' TBS_REVERSED has no effect on the control; it is simply a flag that can be checked.
    If PropReversed = True Then
        If Not (dwStyle And TBS_REVERSED) = TBS_REVERSED Then dwStyle = dwStyle Or TBS_REVERSED
        If Not (dwStyle And TBS_DOWNISLEFT) = TBS_DOWNISLEFT Then dwStyle = dwStyle Or TBS_DOWNISLEFT
    Else
        If (dwStyle And TBS_REVERSED) = TBS_REVERSED Then dwStyle = dwStyle And Not TBS_REVERSED
        If (dwStyle And TBS_DOWNISLEFT) = TBS_DOWNISLEFT Then dwStyle = dwStyle And Not TBS_DOWNISLEFT
    End If
    SetWindowLong SliderHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "Reversed"
End Property

Public Property Get DrawMode() As SldDrawModeConstants
Attribute DrawMode.VB_Description = "Returns/sets a value indicating whether your code or the operating system will handle drawing of the elements."
DrawMode = PropDrawMode
End Property

Public Property Let DrawMode(ByVal Value As SldDrawModeConstants)
Select Case Value
    Case SldDrawModeNormal, SldDrawModeOwnerDraw
        PropDrawMode = Value
    Case Else
        Err.Raise 380
End Select
If SliderHandle <> 0 Then Call ReCreateSlider
UserControl.PropertyChanged "DrawMode"
End Property

Private Sub CreateSlider()
If SliderHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or TBS_AUTOTICKS
If SliderDesignMode = True And PropDrawMode = SldDrawModeOwnerDraw Then
    ' To avoid subclassing the UserControl at design-time just hide the window to visualize unhandled ownerdraw.
    dwStyle = dwStyle And Not WS_VISIBLE
End If
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If PropOrientation = SldOrientationHorizontal Then
    dwStyle = dwStyle Or TBS_HORZ
ElseIf PropOrientation = SldOrientationVertical Then
    dwStyle = dwStyle Or TBS_VERT
End If
Select Case PropTickStyle
    Case SldTickStyleBottomRight
        If PropOrientation = SldOrientationHorizontal Then
            dwStyle = dwStyle Or TBS_BOTTOM
        ElseIf PropOrientation = SldOrientationVertical Then
            dwStyle = dwStyle Or TBS_RIGHT
        End If
    Case SldTickStyleTopLeft
        If PropOrientation = SldOrientationHorizontal Then
            dwStyle = dwStyle Or TBS_TOP
        ElseIf PropOrientation = SldOrientationVertical Then
            dwStyle = dwStyle Or TBS_LEFT
        End If
    Case SldTickStyleBoth
        dwStyle = dwStyle Or TBS_BOTH
    Case SldTickStyleNone
        dwStyle = dwStyle Or TBS_NOTICKS
End Select
If PropShowTip = True Then dwStyle = dwStyle Or TBS_TOOLTIPS
If PropSelectRange = True Then dwStyle = dwStyle Or TBS_ENABLESELRANGE
If PropHideThumb = True Then dwStyle = dwStyle Or TBS_NOTHUMB
If PropReversed = True Then dwStyle = dwStyle Or TBS_REVERSED Or TBS_DOWNISLEFT
If SliderDesignMode = False Then
    ' The WM_NOTIFYFORMAT notification must be handled, which will be sent on control creation.
    ' Thus it is necessary to subclass the parent before the control is created.
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
SliderHandle = CreateWindowEx(dwExStyle, StrPtr("msctls_trackbar32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If SliderHandle <> 0 Then
    SliderToolTipHandle = SendMessage(SliderHandle, TBM_GETTOOLTIPS, 0, ByVal 0&)
    If SliderToolTipHandle <> 0 Then Call ComCtlsInitToolTip(SliderToolTipHandle)
    SendMessage SliderHandle, TBM_SETRANGEMIN, 0, ByVal PropMin
    SendMessage SliderHandle, TBM_SETRANGEMAX, 1, ByVal PropMax
End If
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Value = PropValue
Me.TickFrequency = PropTickFrequency
Me.SmallChange = PropSmallChange
Me.LargeChange = PropLargeChange
Me.TipSide = PropTipSide
If PropSelectRange = True Then Me.SelStart = PropSelStart
If SliderDesignMode = False Then
    If SliderHandle <> 0 Then Call ComCtlsSetSubclass(SliderHandle, Me, 1)
End If
End Sub

Private Sub ReCreateSlider()
If SliderDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Call DestroySlider
    Call CreateSlider
    Call UserControl_Resize
    If Locked = True Then LockWindowUpdate 0
    Me.Refresh
Else
    Call DestroySlider
    Call CreateSlider
    Call UserControl_Resize
End If
End Sub

Private Sub DestroySlider()
If SliderHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(SliderHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow SliderHandle, SW_HIDE
SetParent SliderHandle, 0
DestroyWindow SliderHandle
SliderHandle = 0
SliderToolTipHandle = 0
If SliderTransparentBrush <> 0 Then
    DeleteObject SliderTransparentBrush
    SliderTransparentBrush = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
If SliderTransparentBrush <> 0 Then
    DeleteObject SliderTransparentBrush
    SliderTransparentBrush = 0
End If
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub ClearSel()
Attribute ClearSel.VB_Description = "Clears the current selection range."
If SliderHandle <> 0 Then SendMessage SliderHandle, TBM_CLEARSEL, 1, ByVal 0&
End Sub

Public Function GetNumTicks() As Long
Attribute GetNumTicks.VB_Description = "Returns the number of ticks."
If SliderHandle <> 0 Then GetNumTicks = SendMessage(SliderHandle, TBM_GETNUMTICS, 0, ByVal 0&)
End Function

Public Function GetTickPosition(ByVal Index As Long) As Single
Attribute GetTickPosition.VB_Description = "Returns the current physical position of a tick mark."
If Index < 1 Then Err.Raise 380
If SliderHandle <> 0 Then
    Dim RetVal As Long
    RetVal = SendMessage(SliderHandle, TBM_GETTICPOS, Index - 1, ByVal 0&)
    If RetVal > -1 Then
        If PropOrientation = SldOrientationHorizontal Then
            GetTickPosition = UserControl.ScaleX(RetVal, vbPixels, vbContainerPosition)
        ElseIf PropOrientation = SldOrientationVertical Then
            GetTickPosition = UserControl.ScaleY(RetVal, vbPixels, vbContainerPosition)
        End If
    Else
        Err.Raise 380
    End If
End If
End Function

Public Property Get ThumbLeft() As Single
Attribute ThumbLeft.VB_Description = "Returns the left coordinate of the thumb marker."
Attribute ThumbLeft.VB_MemberFlags = "400"
If SliderHandle <> 0 Then
    Dim RC As RECT
    SendMessage SliderHandle, TBM_GETTHUMBRECT, 0, ByVal VarPtr(RC)
    ThumbLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
End If
End Property

Public Property Get ThumbTop() As Single
Attribute ThumbTop.VB_Description = "Returns the top coordinate of the thumb marker."
Attribute ThumbTop.VB_MemberFlags = "400"
If SliderHandle <> 0 Then
    Dim RC As RECT
    SendMessage SliderHandle, TBM_GETTHUMBRECT, 0, ByVal VarPtr(RC)
    ThumbTop = UserControl.ScaleY(RC.Top, vbPixels, vbContainerPosition)
End If
End Property

Public Property Get ThumbWidth() As Single
Attribute ThumbWidth.VB_Description = "Returns the width of the thumb marker."
Attribute ThumbWidth.VB_MemberFlags = "400"
If SliderHandle <> 0 Then
    Dim RC As RECT
    SendMessage SliderHandle, TBM_GETTHUMBRECT, 0, ByVal VarPtr(RC)
    ThumbWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End If
End Property

Public Property Get ThumbHeight() As Single
Attribute ThumbHeight.VB_Description = "Returns the height of the thumb marker."
Attribute ThumbHeight.VB_MemberFlags = "400"
If SliderHandle <> 0 Then
    Dim RC As RECT
    SendMessage SliderHandle, TBM_GETTHUMBRECT, 0, ByVal VarPtr(RC)
    ThumbHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End If
End Property

Public Property Get ChannelLeft() As Single
Attribute ChannelLeft.VB_Description = "Returns the left coordinate of the channel."
Attribute ChannelLeft.VB_MemberFlags = "400"
If SliderHandle <> 0 Then
    Dim RC As RECT
    SendMessage SliderHandle, TBM_GETCHANNELRECT, 0, ByVal VarPtr(RC)
    ChannelLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
End If
End Property

Public Property Get ChannelTop() As Single
Attribute ChannelTop.VB_Description = "Returns the top coordinate of the channel."
Attribute ChannelTop.VB_MemberFlags = "400"
If SliderHandle <> 0 Then
    Dim RC As RECT
    SendMessage SliderHandle, TBM_GETCHANNELRECT, 0, ByVal VarPtr(RC)
    ChannelTop = UserControl.ScaleY(RC.Top, vbPixels, vbContainerPosition)
End If
End Property

Public Property Get ChannelWidth() As Single
Attribute ChannelWidth.VB_Description = "Returns the width of the channel."
Attribute ChannelWidth.VB_MemberFlags = "400"
If SliderHandle <> 0 Then
    Dim RC As RECT
    SendMessage SliderHandle, TBM_GETCHANNELRECT, 0, ByVal VarPtr(RC)
    ChannelWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End If
End Property

Public Property Get ChannelHeight() As Single
Attribute ChannelHeight.VB_Description = "Returns the height of the channel."
Attribute ChannelHeight.VB_MemberFlags = "400"
If SliderHandle <> 0 Then
    Dim RC As RECT
    SendMessage SliderHandle, TBM_GETCHANNELRECT, 0, ByVal VarPtr(RC)
    ChannelHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End If
End Property

Private Sub SetVisualStylesToolTip()
If SliderHandle <> 0 Then
    If SliderToolTipHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles SliderToolTipHandle
        Else
            RemoveVisualStyles SliderToolTipHandle
        End If
    End If
End If
End Sub

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
    Case WM_LBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
    Case WM_MBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
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
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            SliderCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If SliderCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(SliderCharCodeCache And &HFFFF&)
            SliderCharCodeCache = 0
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
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                SliderIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                SliderIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                SliderIsClick = True
            Case WM_MOUSEMOVE
                If SliderMouseOver = False And PropMouseTrack = True Then
                    SliderMouseOver = True
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
                If SliderIsClick = True Then
                    SliderIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If SliderMouseOver = True Then
            SliderMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_VSCROLL, WM_HSCROLL
        If lParam = SliderHandle Then
            Dim RetVal As Long
            RetVal = SendMessage(SliderHandle, TBM_GETPOS, 0, ByVal 0&)
            Select Case LoWord(wParam)
                Case TB_THUMBTRACK, TB_THUMBPOSITION
                    If RetVal <> PropValue Then
                        PropValue = RetVal
                        RaiseEvent Scroll
                    End If
                Case TB_ENDTRACK
                    PropValue = RetVal
                    UserControl.PropertyChanged "Value"
                    On Error Resume Next
                    UserControl.Extender.DataChanged = True
                    On Error GoTo 0
                    RaiseEvent Change
            End Select
        End If
    Case WM_CTLCOLORSTATIC
        WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If PropTransparent = True Then
            SetBkMode wParam, 1
            Dim hDCBmp As Long
            Dim hBmp As Long, hBmpOld As Long
            With UserControl
            If SliderTransparentBrush = 0 Then
                hDCBmp = CreateCompatibleDC(wParam)
                If hDCBmp <> 0 Then
                    hBmp = CreateCompatibleBitmap(wParam, .ScaleWidth, .ScaleHeight)
                    If hBmp <> 0 Then
                        hBmpOld = SelectObject(hDCBmp, hBmp)
                        Dim WndRect As RECT, P1 As POINTAPI
                        GetWindowRect .hWnd, WndRect
                        MapWindowPoints HWND_DESKTOP, .ContainerHwnd, WndRect, 2
                        P1.X = WndRect.Left
                        P1.Y = WndRect.Top
                        SetViewportOrgEx hDCBmp, -P1.X, -P1.Y, P1
                        SendMessage .ContainerHwnd, WM_PAINT, hDCBmp, ByVal 0&
                        SetViewportOrgEx hDCBmp, P1.X, P1.Y, P1
                        SliderTransparentBrush = CreatePatternBrush(hBmp)
                        SelectObject hDCBmp, hBmpOld
                        DeleteObject hBmp
                    End If
                    DeleteDC hDCBmp
                End If
            End If
            End With
            If SliderTransparentBrush <> 0 Then WindowProcUserControl = SliderTransparentBrush
        End If
        Exit Function
    Case WM_CONTEXTMENU
        If wParam = SliderHandle Then
            Dim P2 As POINTAPI
            P2.X = Get_X_lParam(lParam)
            P2.Y = Get_Y_lParam(lParam)
            If P2.X = -1 And P2.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(-1, -1)
            Else
                ScreenToClient SliderHandle, P2
                RaiseEvent ContextMenu(UserControl.ScaleX(P2.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P2.Y, vbPixels, vbContainerPosition))
            End If
        End If
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = SliderHandle Then
            Select Case NM.Code
                Case NM_CUSTOMDRAW
                    Dim NMCD As NMCUSTOMDRAW
                    CopyMemory NMCD, ByVal lParam, LenB(NMCD)
                    Select Case NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            If PropDrawMode = SldDrawModeOwnerDraw Then
                                WindowProcUserControl = CDRF_NOTIFYITEMDRAW
                            Else
                                WindowProcUserControl = CDRF_DODEFAULT
                            End If
                            Exit Function
                        Case CDDS_ITEMPREPAINT
                            If PropDrawMode = SldDrawModeOwnerDraw Then
                                Dim Cancel As Boolean
                                RaiseEvent ItemDraw(NMCD.dwItemSpec, Cancel, NMCD.uItemState, NMCD.hDC, NMCD.RC.Left, NMCD.RC.Top, NMCD.RC.Right, NMCD.RC.Bottom)
                                If Cancel = False Then WindowProcUserControl = CDRF_SKIPDEFAULT Else WindowProcUserControl = CDRF_DODEFAULT
                            Else
                                WindowProcUserControl = CDRF_DODEFAULT
                            End If
                            Exit Function
                    End Select
            End Select
        ElseIf NM.hWndFrom = SliderToolTipHandle And SliderToolTipHandle <> 0 Then
            Select Case NM.Code
                Case TTN_GETDISPINFO
                    Dim NMTTDI As NMTTDISPINFO
                    CopyMemory NMTTDI, ByVal lParam, LenB(NMTTDI)
                    With NMTTDI
                    If PropRightToLeft = True And PropRightToLeftLayout = False Then
                        If Not (.uFlags And TTF_RTLREADING) = TTF_RTLREADING Then
                            .uFlags = .uFlags Or TTF_RTLREADING
                            CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                        End If
                    End If
                    Dim Text As String, Length As Long, OldText As String
                    If .lpszText <> 0 Then Length = lstrlen(.lpszText)
                    If Length > 0 Then
                        Text = String(Length, vbNullChar)
                        CopyMemory ByVal StrPtr(Text), ByVal .lpszText, Length * 2
                    Else
                        Text = Left$(.szText(), InStr(.szText(), vbNullChar) - 1)
                    End If
                    OldText = Text
                    RaiseEvent ModifyTipText(Text)
                    If StrComp(Text, OldText) <> 0 Then
                        With NMTTDI
                        If Len(Text) <= 80 Then
                            Text = Left$(Text & vbNullChar, 80)
                            CopyMemory .szText(0), ByVal StrPtr(Text), LenB(Text)
                        Else
                            Erase .szText()
                        End If
                        .lpszText = StrPtr(Text) ' Apparently the string address must be always set.
                        .hInst = 0
                        End With
                        CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                    End If
                    End With
            End Select
        End If
    Case WM_NOTIFYFORMAT
        Const NF_QUERY As Long = 3
        If lParam = NF_QUERY Then
            Const NFR_UNICODE As Long = 2
            Const NFR_ANSI As Long = 1
            WindowProcUserControl = NFR_UNICODE
            Exit Function
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI SliderHandle
End Function
