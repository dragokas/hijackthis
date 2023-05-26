VERSION 5.00
Begin VB.UserControl SpinBox 
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   PropertyPages   =   "SpinBox.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "SpinBox.ctx":0048
End
Attribute VB_Name = "SpinBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private SpbNumberStyleDecimal, SpbNumberStyleHexadecimal
#End If
Public Enum SpbNumberStyleConstants
SpbNumberStyleDecimal = 0
SpbNumberStyleHexadecimal = 1
End Enum
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
Private Type UDACCEL
nSec As Long
nInc As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMUPDOWN
hdr As NMHDR
iPos As Long
iDelta As Long
End Type
Public Event DownClick()
Attribute DownClick.VB_Description = "Occurs when the position has changed by a down click."
Public Event UpClick()
Attribute UpClick.VB_Description = "Occurs when the position has changed by an up click."
Public Event BeforeChange(ByVal Value As Long, ByRef Delta As Long)
Attribute BeforeChange.VB_Description = "Occurs when the position is about to change."
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event TextChange()
Attribute TextChange.VB_Description = "Occurs when the contents of a control have changed."
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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const ICC_UPDOWN_CLASS As Long = &H10
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
Private Const TME_LEAVE As Long = &H2, TME_NONCLIENT As Long = &H10
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
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
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_NCMOUSELEAVE As Long = &H2A2
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const EM_SETREADONLY As Long = &HCF, ES_READONLY As Long = &H800
Private Const EM_GETSEL As Long = &HB0
Private Const EM_SETSEL As Long = &HB1
Private Const EM_REPLACESEL As Long = &HC2
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const ES_NUMBER As Long = &H2000
Private Const ES_NOHIDESEL As Long = &H100
Private Const ES_LEFT As Long = &H0
Private Const ES_CENTER As Long = &H1
Private Const ES_RIGHT As Long = &H2
Private Const UDN_FIRST As Long = (-721)
Private Const UDN_DELTAPOS As Long = (UDN_FIRST - 1)
Private Const UDS_WRAP As Long = &H1
Private Const UDS_SETBUDDYINT As Long = &H2
Private Const UDS_ALIGNRIGHT As Long = &H4
Private Const UDS_ALIGNLEFT As Long = &H8
Private Const UDS_ARROWKEYS As Long = &H20
Private Const UDS_NOTHOUSANDS As Long = &H80
Private Const UDS_HOTTRACK As Long = &H100
Private Const WM_USER As Long = &H400
Private Const UM_CHECKVALUE As Long = (WM_USER + 300)
Private Const UDM_SETRANGE As Long = (WM_USER + 101) ' 16 bit
Private Const UDM_GETRANGE As Long = (WM_USER + 102) ' 16 bit
Private Const UDM_SETRANGE32 As Long = (WM_USER + 111)
Private Const UDM_GETRANGE32 As Long = (WM_USER + 112)
Private Const UDM_SETPOS As Long = (WM_USER + 103) ' 16 bit
Private Const UDM_GETPOS As Long = (WM_USER + 104) ' 16 bit
Private Const UDM_SETPOS32 As Long = (WM_USER + 113)
Private Const UDM_GETPOS32 As Long = (WM_USER + 114)
Private Const UDM_SETBUDDY As Long = (WM_USER + 105)
Private Const UDM_GETBUDDY As Long = (WM_USER + 106)
Private Const UDM_SETACCEL As Long = (WM_USER + 107)
Private Const UDM_GETACCEL As Long = (WM_USER + 108)
Private Const UDM_SETBASE As Long = (WM_USER + 109)
Private Const UDM_GETBASE As Long = (WM_USER + 110)
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETUNICODEFORMAT As Long = (CCM_FIRST + 5)
Private Const UDM_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private SpinBoxUpDownHandle As Long, SpinBoxEditHandle As Long
Private SpinBoxFontHandle As Long
Private SpinBoxCharCodeCache As Long
Private SpinBoxMouseOver(0 To 2) As Boolean
Private SpinBoxDesignMode As Boolean
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropMin As Long, PropMax As Long
Private PropValue As Long, PropIncrement As Long
Private PropWrap As Boolean
Private PropHotTracking As Boolean
Private PropAlignment As CCLeftRightAlignmentConstants
Private PropThousandsSeparator As Boolean
Private PropNumberStyle As SpbNumberStyleConstants
Private PropArrowKeysChange As Boolean
Private PropAllowOnlyNumbers As Boolean
Private PropTextAlignment As VBRUN.AlignmentConstants
Private PropLocked As Boolean
Private PropHideSelection As Boolean

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
Call ComCtlsInitCC(ICC_STANDARD_CLASSES Or ICC_UPDOWN_CLASS)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
SpinBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropMin = 0
PropMax = 100
PropValue = 0
PropIncrement = 1
PropWrap = False
PropHotTracking = True
If PropRightToLeft = False Then PropAlignment = CCLeftRightAlignmentRight Else PropAlignment = CCLeftRightAlignmentLeft
PropThousandsSeparator = True
PropNumberStyle = SpbNumberStyleDecimal
PropArrowKeysChange = True
PropAllowOnlyNumbers = False
If PropRightToLeft = False Then PropTextAlignment = vbLeftJustify Else PropTextAlignment = vbRightJustify
PropLocked = False
PropHideSelection = True
Call CreateSpinBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
SpinBoxDesignMode = Not Ambient.UserMode
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
PropMin = .ReadProperty("Min", 0)
PropMax = .ReadProperty("Max", 100)
PropValue = .ReadProperty("Value", 0)
PropIncrement = .ReadProperty("Increment", 1)
PropWrap = .ReadProperty("Wrap", False)
PropHotTracking = .ReadProperty("HotTracking", True)
PropAlignment = .ReadProperty("Alignment", CCLeftRightAlignmentRight)
PropThousandsSeparator = .ReadProperty("ThousandsSeparator", True)
PropNumberStyle = .ReadProperty("NumberStyle", SpbNumberStyleDecimal)
PropArrowKeysChange = .ReadProperty("ArrowKeysChange", True)
PropAllowOnlyNumbers = .ReadProperty("AllowOnlyNumbers", False)
PropTextAlignment = .ReadProperty("TextAlignment", vbLeftJustify)
PropLocked = .ReadProperty("Locked", False)
PropHideSelection = .ReadProperty("HideSelection", True)
End With
Call CreateSpinBox
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
.WriteProperty "Min", PropMin, 0
.WriteProperty "Max", PropMax, 100
.WriteProperty "Value", PropValue, 0
.WriteProperty "Increment", PropIncrement, 1
.WriteProperty "Wrap", PropWrap, False
.WriteProperty "HotTracking", PropHotTracking, True
.WriteProperty "Alignment", PropAlignment, CCLeftRightAlignmentRight
.WriteProperty "ThousandsSeparator", PropThousandsSeparator, True
.WriteProperty "NumberStyle", PropNumberStyle, SpbNumberStyleDecimal
.WriteProperty "ArrowKeysChange", PropArrowKeysChange, True
.WriteProperty "AllowOnlyNumbers", PropAllowOnlyNumbers, False
.WriteProperty "TextAlignment", PropTextAlignment, vbLeftJustify
.WriteProperty "Locked", PropLocked, False
.WriteProperty "HideSelection", PropHideSelection, True
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
If SpinBoxEditHandle <> 0 Then MoveWindow SpinBoxEditHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
If SpinBoxUpDownHandle <> 0 Then SendMessage SpinBoxUpDownHandle, UDM_SETBUDDY, SpinBoxEditHandle, ByVal 0&
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroySpinBox
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
hWnd = SpinBoxUpDownHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndEdit() As Long
Attribute hWndEdit.VB_Description = "Returns a handle to a control."
hWndEdit = SpinBoxEditHandle
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
OldFontHandle = SpinBoxFontHandle
SpinBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If SpinBoxEditHandle <> 0 Then SendMessage SpinBoxEditHandle, WM_SETFONT, SpinBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = SpinBoxFontHandle
SpinBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If SpinBoxUpDownHandle <> 0 Then SendMessage SpinBoxUpDownHandle, WM_SETFONT, SpinBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If SpinBoxUpDownHandle <> 0 And SpinBoxEditHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles SpinBoxUpDownHandle
        ActivateVisualStyles SpinBoxEditHandle
    Else
        RemoveVisualStyles SpinBoxUpDownHandle
        RemoveVisualStyles SpinBoxEditHandle
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
If SpinBoxUpDownHandle <> 0 Then
    EnableWindow SpinBoxUpDownHandle, IIf(Value = True, 1, 0)
    If SpinBoxEditHandle <> 0 Then EnableWindow SpinBoxEditHandle, IIf(Value = True, 1, 0)
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
If SpinBoxDesignMode = False Then Call RefreshMousePointer
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
        If SpinBoxDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If SpinBoxDesignMode = False Then Call RefreshMousePointer
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
If PropRightToLeft = True Then dwMask = WS_EX_RTLREADING
If SpinBoxEditHandle <> 0 Then
    Call ComCtlsSetRightToLeft(SpinBoxEditHandle, dwMask)
    If PropRightToLeft = False Then
        If PropAlignment = CCLeftRightAlignmentLeft Then Me.Alignment = CCLeftRightAlignmentRight
        If PropTextAlignment = vbRightJustify Then Me.TextAlignment = vbLeftJustify
    Else
        If PropAlignment = CCLeftRightAlignmentRight Then Me.Alignment = CCLeftRightAlignmentLeft
        If PropTextAlignment = vbLeftJustify Then Me.TextAlignment = vbRightJustify
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

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum value."
If SpinBoxUpDownHandle <> 0 Then
    SendMessage SpinBoxUpDownHandle, UDM_GETRANGE32, VarPtr(Min), ByVal 0&
Else
    Min = PropMin
End If
End Property

Public Property Let Min(ByVal Value As Long)
If Value <= Me.Max Then
    PropMin = Value
    If Me.Value < PropMin Then Me.Value = PropMin
Else
    If SpinBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If SpinBoxUpDownHandle <> 0 Then SendMessage SpinBoxUpDownHandle, UDM_SETRANGE32, PropMin, ByVal PropMax
Me.Refresh
UserControl.PropertyChanged "Min"
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum value."
If SpinBoxUpDownHandle <> 0 Then
    SendMessage SpinBoxUpDownHandle, UDM_GETRANGE32, 0, ByVal VarPtr(Max)
Else
    Max = PropMax
End If
End Property

Public Property Let Max(ByVal Value As Long)
If Value >= Me.Min Then
    PropMax = Value
    If Me.Value > PropMax Then Me.Value = PropMax
Else
    If SpinBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If SpinBoxUpDownHandle <> 0 Then SendMessage SpinBoxUpDownHandle, UDM_SETRANGE32, PropMin, ByVal PropMax
Me.Refresh
UserControl.PropertyChanged "Max"
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the current position."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "123c"
If SpinBoxUpDownHandle <> 0 Then
    Value = SendMessage(SpinBoxUpDownHandle, UDM_GETPOS32, 0, ByVal 0&)
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
If SpinBoxUpDownHandle <> 0 Then SendMessage SpinBoxUpDownHandle, UDM_SETPOS32, 0, ByVal PropValue
UserControl.PropertyChanged "Value"
If Changed = True Then
    On Error Resume Next
    UserControl.Extender.DataChanged = True
    On Error GoTo 0
    RaiseEvent Change
End If
End Property

Public Property Get Increment() As Long
Attribute Increment.VB_Description = "Returns/sets the position change increment."
If SpinBoxUpDownHandle <> 0 Then
    Dim Accel As UDACCEL
    SendMessage SpinBoxUpDownHandle, UDM_GETACCEL, 1, Accel
    Increment = Accel.nInc
Else
    Increment = PropIncrement
End If
End Property

Public Property Let Increment(ByVal Value As Long)
PropIncrement = Value
If SpinBoxUpDownHandle <> 0 Then
    Dim Accel As UDACCEL
    Accel.nSec = 0
    Accel.nInc = PropIncrement
    SendMessage SpinBoxUpDownHandle, UDM_SETACCEL, 1, Accel
End If
UserControl.PropertyChanged "Increment"
End Property

Public Property Get Wrap() As Boolean
Attribute Wrap.VB_Description = "Returns/sets a value that determines whether or not the position will be wrapped if it is incremented or decremented beyond the ending or beginning of the range."
Wrap = PropWrap
End Property

Public Property Let Wrap(ByVal Value As Boolean)
PropWrap = Value
If SpinBoxUpDownHandle <> 0 Then Call ReCreateSpinBox
UserControl.PropertyChanged "Wrap"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets a value that determines whether or not the control highlights the up arrow and down arrow as the pointer passes over them. This flag is ignored on Windows XP (or above) when the desktop theme overrides it."
HotTracking = PropHotTracking
End Property

Public Property Let HotTracking(ByVal Value As Boolean)
PropHotTracking = Value
If SpinBoxUpDownHandle <> 0 Then Call ReCreateSpinBox
UserControl.PropertyChanged "HotTracking"
End Property

Public Property Get Alignment() As CCLeftRightAlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment."
Alignment = PropAlignment
End Property

Public Property Let Alignment(ByVal Value As CCLeftRightAlignmentConstants)
Select Case Value
    Case CCLeftRightAlignmentLeft, CCLeftRightAlignmentRight
        PropAlignment = Value
    Case Else
        Err.Raise 380
End Select
If SpinBoxUpDownHandle <> 0 Then Call ReCreateSpinBox
UserControl.PropertyChanged "Alignment"
End Property

Public Property Get ThousandsSeparator() As Boolean
Attribute ThousandsSeparator.VB_Description = "Returns/sets a value that determines whether a thousand separator will be insert between every three decimal digits or not."
ThousandsSeparator = PropThousandsSeparator
End Property

Public Property Let ThousandsSeparator(ByVal Value As Boolean)
PropThousandsSeparator = Value
If SpinBoxUpDownHandle <> 0 Then Call ReCreateSpinBox
UserControl.PropertyChanged "ThousandsSeparator"
End Property

Public Property Get NumberStyle() As SpbNumberStyleConstants
Attribute NumberStyle.VB_Description = "Returns/sets the number style."
If SpinBoxUpDownHandle <> 0 Then
    Select Case SendMessage(SpinBoxUpDownHandle, UDM_GETBASE, 0, ByVal 0&)
        Case 10
            NumberStyle = SpbNumberStyleDecimal
        Case 16
            NumberStyle = SpbNumberStyleHexadecimal
        Case Else
            NumberStyle = PropNumberStyle
    End Select
Else
    NumberStyle = PropNumberStyle
End If
End Property

Public Property Let NumberStyle(ByVal Value As SpbNumberStyleConstants)
Select Case Value
    Case SpbNumberStyleDecimal, SpbNumberStyleHexadecimal
        PropNumberStyle = Value
    Case Else
        Err.Raise 380
End Select
If SpinBoxUpDownHandle <> 0 Then
    Select Case PropNumberStyle
        Case SpbNumberStyleDecimal
            SendMessage SpinBoxUpDownHandle, UDM_SETBASE, 10, ByVal 0&
        Case SpbNumberStyleHexadecimal
            SendMessage SpinBoxUpDownHandle, UDM_SETBASE, 16, ByVal 0&
    End Select
End If
UserControl.PropertyChanged "NumberStyle"
End Property

Public Property Get ArrowKeysChange() As Boolean
Attribute ArrowKeysChange.VB_Description = "Returns/sets a value that determines whether or not the position can be incremented and decrement when the up arrow and down arrow keys are pressed."
ArrowKeysChange = PropArrowKeysChange
End Property

Public Property Let ArrowKeysChange(ByVal Value As Boolean)
PropArrowKeysChange = Value
If SpinBoxUpDownHandle <> 0 Then Call ReCreateSpinBox
UserControl.PropertyChanged "ArrowKeysChange"
End Property

Public Property Get AllowOnlyNumbers() As Boolean
Attribute AllowOnlyNumbers.VB_Description = "Returns/sets a value indicating if only numbers are allowed to be entered."
AllowOnlyNumbers = PropAllowOnlyNumbers
End Property

Public Property Let AllowOnlyNumbers(ByVal Value As Boolean)
PropAllowOnlyNumbers = Value
If SpinBoxEditHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(SpinBoxEditHandle, GWL_STYLE)
    If PropAllowOnlyNumbers = True Then
        If Not (dwStyle And ES_NUMBER) = ES_NUMBER Then dwStyle = dwStyle Or ES_NUMBER
    Else
        If (dwStyle And ES_NUMBER) = ES_NUMBER Then dwStyle = dwStyle And Not ES_NUMBER
    End If
    SetWindowLong SpinBoxEditHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "AllowOnlyNumbers"
End Property

Public Property Get TextAlignment() As VBRUN.AlignmentConstants
Attribute TextAlignment.VB_Description = "Returns/sets the alignment of the displayed text."
TextAlignment = PropTextAlignment
End Property

Public Property Let TextAlignment(ByVal Value As VBRUN.AlignmentConstants)
Select Case Value
    Case vbLeftJustify, vbCenter, vbRightJustify
        PropTextAlignment = Value
    Case Else
        Err.Raise 380
End Select
If SpinBoxEditHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(SpinBoxEditHandle, GWL_STYLE)
    If (dwStyle And ES_LEFT) = ES_LEFT Then dwStyle = dwStyle And Not ES_LEFT
    If (dwStyle And ES_CENTER) = ES_CENTER Then dwStyle = dwStyle And Not ES_CENTER
    If (dwStyle And ES_RIGHT) = ES_RIGHT Then dwStyle = dwStyle And Not ES_RIGHT
    Select Case PropTextAlignment
        Case vbLeftJustify
            dwStyle = dwStyle Or ES_LEFT
        Case vbCenter
            dwStyle = dwStyle Or ES_CENTER
        Case vbRightJustify
            dwStyle = dwStyle Or ES_RIGHT
    End Select
    SetWindowLong SpinBoxEditHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "TextAlignment"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents can be edited."
If SpinBoxEditHandle <> 0 Then
    Locked = CBool((GetWindowLong(SpinBoxEditHandle, GWL_STYLE) And ES_READONLY) <> 0)
Else
    Locked = PropLocked
End If
End Property

Public Property Let Locked(ByVal Value As Boolean)
PropLocked = Value
If SpinBoxEditHandle <> 0 Then SendMessage SpinBoxEditHandle, EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Locked"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value indicating if the selection in an edit control is hidden when the control loses focus."
HideSelection = PropHideSelection
End Property

Public Property Let HideSelection(ByVal Value As Boolean)
PropHideSelection = Value
If SpinBoxUpDownHandle <> 0 Then Call ReCreateSpinBox
UserControl.PropertyChanged "HideSelection"
End Property

Private Sub CreateSpinBox()
If SpinBoxUpDownHandle <> 0 Or SpinBoxEditHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwStyleEdit As Long, dwExStyleEdit As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or UDS_SETBUDDYINT
dwStyleEdit = WS_CHILD Or WS_VISIBLE Or ES_AUTOHSCROLL
dwExStyleEdit = WS_EX_CLIENTEDGE
If PropRightToLeft = True Then dwExStyleEdit = dwExStyleEdit Or WS_EX_RTLREADING
If PropWrap = True Then dwStyle = dwStyle Or UDS_WRAP
If PropHotTracking = True Then dwStyle = dwStyle Or UDS_HOTTRACK
Select Case PropAlignment
    Case CCLeftRightAlignmentLeft
        dwStyle = dwStyle Or UDS_ALIGNLEFT
    Case CCLeftRightAlignmentRight
        dwStyle = dwStyle Or UDS_ALIGNRIGHT
End Select
If PropThousandsSeparator = False Then dwStyle = dwStyle Or UDS_NOTHOUSANDS
If PropArrowKeysChange = True Then dwStyle = dwStyle Or UDS_ARROWKEYS
If PropAllowOnlyNumbers = True Then dwStyleEdit = dwStyleEdit Or ES_NUMBER
Select Case PropTextAlignment
    Case vbLeftJustify
        dwStyleEdit = dwStyleEdit Or ES_LEFT
    Case vbCenter
        dwStyleEdit = dwStyleEdit Or ES_CENTER
    Case vbRightJustify
        dwStyleEdit = dwStyleEdit Or ES_RIGHT
End Select
If PropLocked = True Then dwStyleEdit = dwStyleEdit Or ES_READONLY
If PropHideSelection = False Then dwStyleEdit = dwStyleEdit Or ES_NOHIDESEL
SpinBoxEditHandle = CreateWindowEx(dwExStyleEdit, StrPtr("Edit"), 0, dwStyleEdit, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If SpinBoxEditHandle <> 0 Then
    SpinBoxUpDownHandle = CreateWindowEx(0, StrPtr("msctls_updown32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
    If SpinBoxUpDownHandle <> 0 Then
        SendMessage SpinBoxUpDownHandle, UDM_SETUNICODEFORMAT, 1, ByVal 0&
        SendMessage SpinBoxUpDownHandle, UDM_SETRANGE32, PropMin, ByVal PropMax
        SendMessage SpinBoxUpDownHandle, UDM_SETBUDDY, SpinBoxEditHandle, ByVal 0&
        If PropNumberStyle = SpbNumberStyleHexadecimal Then SendMessage SpinBoxUpDownHandle, UDM_SETBASE, 16, ByVal 0&
    End If
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Value = PropValue
Me.Increment = PropIncrement
If SpinBoxDesignMode = False Then
    If SpinBoxUpDownHandle <> 0 Then Call ComCtlsSetSubclass(SpinBoxUpDownHandle, Me, 1)
    If SpinBoxEditHandle <> 0 Then Call ComCtlsSetSubclass(SpinBoxEditHandle, Me, 2)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
End If
End Sub

Private Sub ReCreateSpinBox()
If SpinBoxDesignMode = False Then
    Dim Locked As Boolean
    With Me
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Dim Text As String, SelStart As Long, SelEnd As Long
    Text = .Text
    If SpinBoxEditHandle <> 0 Then SendMessage SpinBoxEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    Call DestroySpinBox
    Call CreateSpinBox
    Call UserControl_Resize
    .Text = Text
    If SpinBoxEditHandle <> 0 Then SendMessage SpinBoxEditHandle, EM_SETSEL, SelStart, ByVal SelEnd
    If Locked = True Then LockWindowUpdate 0
    .Refresh
    End With
Else
    Call DestroySpinBox
    Call CreateSpinBox
    Call UserControl_Resize
End If
End Sub

Private Sub DestroySpinBox()
If SpinBoxUpDownHandle = 0 Or SpinBoxEditHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(SpinBoxUpDownHandle)
Call ComCtlsRemoveSubclass(SpinBoxEditHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow SpinBoxUpDownHandle, SW_HIDE
ShowWindow SpinBoxEditHandle, SW_HIDE
SendMessage SpinBoxUpDownHandle, UDM_SETBUDDY, 0, ByVal 0&
SetParent SpinBoxUpDownHandle, 0
SetParent SpinBoxEditHandle, 0
DestroyWindow SpinBoxUpDownHandle
DestroyWindow SpinBoxEditHandle
SpinBoxUpDownHandle = 0
SpinBoxEditHandle = 0
If SpinBoxFontHandle <> 0 Then
    DeleteObject SpinBoxFontHandle
    SpinBoxFontHandle = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub SetAcceleration(ByVal Delays As Variant, ByVal Increments As Variant)
Attribute SetAcceleration.VB_Description = "Method to set an acceleration. The delays array specify the amount of time to elapse (in seconds) before the position change increment specified in the increments array is used."
If SpinBoxUpDownHandle <> 0 Then
    If IsArray(Delays) And IsArray(Increments) Then
        Dim Ptr(0 To 1) As Long
        CopyMemory Ptr(0), ByVal UnsignedAdd(VarPtr(Delays), 8), 4
        CopyMemory Ptr(1), ByVal UnsignedAdd(VarPtr(Increments), 8), 4
        If Ptr(0) <> 0 And Ptr(1) <> 0 Then
            Dim DimensionCount(0 To 1) As Integer
            CopyMemory DimensionCount(0), ByVal Ptr(0), 2
            CopyMemory DimensionCount(1), ByVal Ptr(1), 2
            If DimensionCount(0) = 1 And DimensionCount(1) = 1 Then
                If LBound(Delays) = LBound(Increments) And UBound(Delays) = UBound(Increments) Then
                    Dim AccelArr() As UDACCEL, Count As Long, i As Long
                    For i = LBound(Delays) To UBound(Delays)
                        Select Case VarType(Delays(i))
                            Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
                                ReDim Preserve AccelArr(0 To Count) As UDACCEL
                                AccelArr(Count).nSec = CLng(Delays(i))
                                Select Case VarType(Increments(i))
                                    Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
                                        AccelArr(Count).nInc = CLng(Increments(i))
                                End Select
                                Count = Count + 1
                        End Select
                    Next i
                    If Count > 0 Then
                        SendMessage SpinBoxUpDownHandle, UDM_SETACCEL, Count, ByVal VarPtr(AccelArr(0))
                    Else
                        Me.Increment = PropIncrement
                    End If
                Else
                    Err.Raise Number:=5, Description:="Array boundaries are not equal"
                End If
            Else
                Err.Raise Number:=5, Description:="Array must be single dimensioned"
            End If
        Else
            Err.Raise Number:=91, Description:="Array is not allocated"
        End If
    ElseIf IsEmpty(Delays) Then
        Me.Increment = PropIncrement
    Else
        Err.Raise 380
    End If
End If
End Sub

Public Sub ValidateText()
Attribute ValidateText.VB_Description = "Method that validates and updates the text displayed in the spin box."
If SpinBoxUpDownHandle <> 0 Then
    Dim Text As String, Value As Long
    Text = String(SendMessage(SpinBoxEditHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage SpinBoxEditHandle, WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
    Value = SendMessage(SpinBoxUpDownHandle, UDM_GETPOS32, 0, ByVal 0&)
    If Not Text = CStr(Value) Then
        Text = CStr(Value)
        SendMessage SpinBoxEditHandle, WM_SETTEXT, 0, ByVal StrPtr(Text)
    End If
End If
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_MemberFlags = "400"
If SpinBoxEditHandle <> 0 Then
    Text = String(SendMessage(SpinBoxEditHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage SpinBoxEditHandle, WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
End If
End Property

Public Property Let Text(ByVal Value As String)
If SpinBoxEditHandle <> 0 Then SendMessage SpinBoxEditHandle, WM_SETTEXT, 0, ByVal StrPtr(Value)
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute SelStart.VB_MemberFlags = "400"
If SpinBoxEditHandle <> 0 Then SendMessage SpinBoxEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal 0&
End Property

Public Property Let SelStart(ByVal Value As Long)
If SpinBoxEditHandle <> 0 Then
    If Value >= 0 Then
        SendMessage SpinBoxEditHandle, EM_SETSEL, Value, ByVal Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
If SpinBoxEditHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage SpinBoxEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    SelLength = SelEnd - SelStart
End If
End Property

Public Property Let SelLength(ByVal Value As Long)
If SpinBoxEditHandle <> 0 Then
    If Value >= 0 Then
        Dim SelStart As Long
        SendMessage SpinBoxEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal 0&
        SendMessage SpinBoxEditHandle, EM_SETSEL, SelStart, ByVal SelStart + Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
If SpinBoxEditHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage SpinBoxEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    On Error Resume Next
    SelText = Mid(Me.Text, SelStart + 1, (SelEnd - SelStart))
    On Error GoTo 0
End If
End Property

Public Property Let SelText(ByVal Value As String)
If SpinBoxEditHandle <> 0 Then
    If StrPtr(Value) = 0 Then Value = ""
    SendMessage SpinBoxEditHandle, EM_REPLACESEL, 0, ByVal StrPtr(Value)
End If
End Property

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcEdit(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg = UM_CHECKVALUE Then
    If wParam <> PropValue Then
        PropValue = wParam
        UserControl.PropertyChanged "Value"
        On Error Resume Next
        UserControl.Extender.DataChanged = True
        On Error GoTo 0
        RaiseEvent Change
    End If
    Exit Function
End If
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim P As POINTAPI
        P.X = Get_X_lParam(lParam)
        P.Y = Get_Y_lParam(lParam)
        MapWindowPoints hWnd, UserControl.hWnd, P, 1
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(P.X, vbPixels, vbTwips)
        Y = UserControl.ScaleY(P.Y, vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MOUSEMOVE
                If (SpinBoxMouseOver(0) = False And PropMouseTrack = True) Or (SpinBoxMouseOver(2) = False And PropMouseTrack = True) Then
                    If SpinBoxMouseOver(0) = False And PropMouseTrack = True Then SpinBoxMouseOver(0) = True
                    If SpinBoxMouseOver(2) = False And PropMouseTrack = True Then
                        SpinBoxMouseOver(2) = True
                        RaiseEvent MouseEnter
                    End If
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP
                RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MBUTTONUP
                RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONUP
                RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
        End Select
    Case WM_MOUSELEAVE
        SpinBoxMouseOver(0) = False
        If SpinBoxMouseOver(2) = True Then
            Dim Pos As Long
            Pos = GetMessagePos()
            If WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos)) <> SpinBoxEditHandle Or SpinBoxEditHandle = 0 Then
                SpinBoxMouseOver(2) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function

Private Function WindowProcEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
                WindowProcEdit = 1
                Exit Function
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then
                    SetCursor PropMouseIcon.Handle
                    WindowProcEdit = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_LBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            SpinBoxCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If SpinBoxCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(SpinBoxCharCodeCache And &HFFFF&)
            SpinBoxCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(wParam And &HFFFF&)
        End If
        RaiseEvent KeyPress(KeyChar)
        If (wParam And &HFFFF&) <> 0 And KeyChar = 0 Then
            Exit Function
        Else
            wParam = CIntToUInt(KeyChar)
        End If
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
End Select
WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
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
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MOUSEMOVE
                If (SpinBoxMouseOver(1) = False And PropMouseTrack = True) Or (SpinBoxMouseOver(2) = False And PropMouseTrack = True) Then
                    If SpinBoxMouseOver(1) = False And PropMouseTrack = True Then SpinBoxMouseOver(1) = True
                    If SpinBoxMouseOver(2) = False And PropMouseTrack = True Then
                        SpinBoxMouseOver(2) = True
                        RaiseEvent MouseEnter
                    End If
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP
                RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MBUTTONUP
                RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONUP
                RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
        End Select
    Case WM_MOUSELEAVE
        Dim TME As TRACKMOUSEEVENTSTRUCT
        With TME
        .cbSize = LenB(TME)
        .hWndTrack = hWnd
        .dwFlags = TME_LEAVE Or TME_NONCLIENT
        End With
        TrackMouseEvent TME
    Case WM_NCMOUSELEAVE
        SpinBoxMouseOver(1) = False
        If SpinBoxMouseOver(2) = True Then
            Dim Pos As Long
            Pos = GetMessagePos()
            If WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos)) <> SpinBoxUpDownHandle Or SpinBoxUpDownHandle = 0 Then
                SpinBoxMouseOver(2) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = SpinBoxUpDownHandle Then
            If NM.Code = UDN_DELTAPOS Then
                Dim NMUD As NMUPDOWN
                CopyMemory NMUD, ByVal lParam, LenB(NMUD)
                RaiseEvent BeforeChange(NMUD.iPos, NMUD.iDelta)
                Select Case NMUD.iDelta
                    Case 0
                        WindowProcUserControl = 1
                        Exit Function
                    Case Is < 0
                        RaiseEvent DownClick
                    Case Is > 0
                        RaiseEvent UpClick
                End Select
            End If
        End If
    Case WM_COMMAND
        Static ChangeFrozen As Boolean
        Const EN_UPDATE As Long = &H400, EN_CHANGE As Long = &H300
        Select Case HiWord(wParam)
            Case EN_UPDATE
                If PropAllowOnlyNumbers = True Then
                    If ComCtlsSupportLevel() <= 1 And ChangeFrozen = False Then
                        Dim Text As String
                        Text = String(SendMessage(lParam, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
                        SendMessage lParam, WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
                        If Not Text = vbNullString Then
                            On Error Resume Next
                            If Left(Text, 2) = "0x" And PropNumberStyle = SpbNumberStyleHexadecimal Then
                                Text = CStr(CLng("&H" & Mid(Text, 3)))
                            Else
                                Text = CStr(CLng(Text))
                            End If
                            If Err.Number <> 0 Then
                                ChangeFrozen = True
                                SendMessage lParam, WM_SETTEXT, 0, ByVal 0&
                                SendMessage lParam, WM_CHAR, 0, ByVal 0&
                                Exit Function
                            End If
                            On Error GoTo 0
                        End If
                    End If
                End If
            Case EN_CHANGE
                If ChangeFrozen = False Then
                    RaiseEvent TextChange
                    If SpinBoxUpDownHandle <> 0 Then PostMessage SpinBoxUpDownHandle, UM_CHECKVALUE, SendMessage(SpinBoxUpDownHandle, UDM_GETPOS32, 0, ByVal 0&), ByVal 0&
                Else
                    ChangeFrozen = False
                    Exit Function
                End If
        End Select
    Case WM_VSCROLL, WM_HSCROLL
        If lParam = SpinBoxUpDownHandle Then
            Dim NewValue As Long
            NewValue = SendMessage(SpinBoxUpDownHandle, UDM_GETPOS32, 0, ByVal 0&)
            If PropValue <> NewValue Then
                PropValue = NewValue
                UserControl.PropertyChanged "Value"
                On Error Resume Next
                UserControl.Extender.DataChanged = True
                On Error GoTo 0
                RaiseEvent Change
            End If
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI SpinBoxEditHandle
End Function
