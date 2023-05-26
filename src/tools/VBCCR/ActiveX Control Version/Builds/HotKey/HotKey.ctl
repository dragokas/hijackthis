VERSION 5.00
Begin VB.UserControl HotKey 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "HotKey.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "HotKey.ctx":0035
End
Attribute VB_Name = "HotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const ImplementThemedBorder = True

#If False Then
Private HkeInvalidKeyCombinationNone, HkeInvalidKeyCombinationShift, HkeInvalidKeyCombinationCtrl, HkeInvalidKeyCombinationAlt, HkeInvalidKeyCombinationShiftCtrl, HkeInvalidKeyCombinationShiftAlt, HkeInvalidKeyCombinationCtrlAlt, HkeInvalidKeyCombinationShiftCtrlAlt
#End If
Private Const HKCOMB_NONE As Long = &H1
Private Const HKCOMB_S As Long = &H2
Private Const HKCOMB_C As Long = &H4
Private Const HKCOMB_A As Long = &H8
Private Const HKCOMB_SC As Long = &H10
Private Const HKCOMB_SA As Long = &H20
Private Const HKCOMB_CA As Long = &H40
Private Const HKCOMB_SCA As Long = &H80
Public Enum HkeInvalidKeyCombinationConstants
HkeInvalidKeyCombinationNone = HKCOMB_NONE
HkeInvalidKeyCombinationShift = HKCOMB_S
HkeInvalidKeyCombinationCtrl = HKCOMB_C
HkeInvalidKeyCombinationAlt = HKCOMB_A
HkeInvalidKeyCombinationShiftCtrl = HKCOMB_SC
HkeInvalidKeyCombinationShiftAlt = HKCOMB_SA
HkeInvalidKeyCombinationCtrlAlt = HKCOMB_CA
HkeInvalidKeyCombinationShiftCtrlAlt = HKCOMB_SCA
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
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
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
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwThreadID As Long) As Long
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextW" (ByVal lParam As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExW" (ByVal wCode As Long, ByVal wMapType As Long, ByVal hKL As Long) As Long

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
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

#End If

Private Const ICC_HOTKEY_CLASS As Long = &H40
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80, RDW_NOCHILDREN As Long = &H40, RDW_FRAME As Long = &H400
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const DCX_WINDOW As Long = &H1
Private Const DCX_INTERSECTRGN As Long = &H80
Private Const DCX_USESTYLE As Long = &H10000
Private Const GCL_STYLE As Long = (-26)
Private Const CS_DBLCLKS As Long = &H8
Private Const MAPVK_VK_TO_VSC As Long = 0
Private Const HOTKEYF_EXT As Long = &H8
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const SW_HIDE As Long = &H0
Private Const GA_ROOT As Long = 2
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_ENABLE As Long = &HA
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_STYLECHANGED As Long = &H7D
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
Private Const WM_NCPAINT As Long = &H85
Private Const WM_COMMAND As Long = &H111
Private Const WM_SETFONT As Long = &H30
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_SETHOTKEY As Long = &H32
Private Const WM_USER As Long = &H400
Private Const HKM_SETHOTKEY As Long = (WM_USER + 1)
Private Const HKM_GETHOTKEY As Long = (WM_USER + 2)
Private Const HKM_SETRULES As Long = (WM_USER + 3)
Private Const EN_CHANGE As Long = &H300
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private HotKeyHandle As Long
Private HotKeyFontHandle As Long
Private HotKeyBackColorBrush As Long
Private HotKeyCharCodeCache As Long
Private HotKeyIsClick As Boolean
Private HotKeyMouseOver As Boolean
Private HotKeyDesignMode As Boolean
Private HotKeyFocused As Boolean
Private HotKeyEnabledVisualStyles As Boolean
Private HotKeyDblClickSupported As Boolean, HotKeyIsDblClick As Boolean
Private HotKeyDblClickTime As Long, HotKeyDblClickTickCount As Double
Private HotKeyDblClickCX As Long, HotKeyDblClickCY As Long
Private HotKeyDblClickX As Long, HotKeyDblClickY As Long
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropBackColor As OLE_COLOR
Private PropBorderStyle As CCBorderStyleConstants

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
        Case vbKeyReturn, vbKeyTab, vbKeyEscape
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
Call ComCtlsInitCC(ICC_HOTKEY_CLASS)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
HotKeyDblClickTime = GetDoubleClickTime()
Const SM_CXDOUBLECLK As Long = 36
Const SM_CYDOUBLECLK As Long = 37
HotKeyDblClickCX = GetSystemMetrics(SM_CXDOUBLECLK)
HotKeyDblClickCY = GetSystemMetrics(SM_CYDOUBLECLK)
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
HotKeyDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropBackColor = vbWindowBackground
PropBorderStyle = CCBorderStyleSunken
Call CreateHotKey
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
HotKeyDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSunken)
End With
Call CreateHotKey
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSunken
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
If HotKeyHandle <> 0 Then MoveWindow HotKeyHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyHotKey
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
hWnd = HotKeyHandle
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
OldFontHandle = HotKeyFontHandle
HotKeyFontHandle = CreateGDIFontFromOLEFont(PropFont)
If HotKeyHandle <> 0 Then SendMessage HotKeyHandle, WM_SETFONT, HotKeyFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = HotKeyFontHandle
HotKeyFontHandle = CreateGDIFontFromOLEFont(PropFont)
If HotKeyHandle <> 0 Then SendMessage HotKeyHandle, WM_SETFONT, HotKeyFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
HotKeyEnabledVisualStyles = EnabledVisualStyles()
If HotKeyHandle <> 0 And HotKeyEnabledVisualStyles = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles HotKeyHandle
    Else
        RemoveVisualStyles HotKeyHandle
    End If
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
If HotKeyHandle <> 0 Then EnableWindow HotKeyHandle, IIf(Value = True, 1, 0)
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
If HotKeyDesignMode = False Then Call RefreshMousePointer
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
        If HotKeyDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If HotKeyDesignMode = False Then Call RefreshMousePointer
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
If HotKeyHandle <> 0 Then
    If HotKeyBackColorBrush <> 0 Then DeleteObject HotKeyBackColorBrush
    HotKeyBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
End If
Me.Refresh

#If ImplementThemedBorder = True Then

If PropBorderStyle = CCBorderStyleSunken Then
    ' Redraw the border to consider the new back color for the themed border, if any.
    RedrawWindow UserControl.hWnd, 0, 0, RDW_FRAME Or RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_NOCHILDREN
End If

#End If

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
If HotKeyHandle <> 0 Then Call ComCtlsChangeBorderStyle(HotKeyHandle, PropBorderStyle)
UserControl.PropertyChanged "BorderStyle"
End Property

Private Sub CreateHotKey()
If HotKeyHandle <> 0 Then Exit Sub
Dim dwStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
HotKeyHandle = CreateWindowEx(0, StrPtr("msctls_hotkey32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If PropBorderStyle <> CCBorderStyleSunken Then
    ' According to MSDN:
    ' WS_EX_CLIENTEDGE is predefined when control receives WM_NCCREATE.
    Me.BorderStyle = PropBorderStyle
End If
If HotKeyDesignMode = False Then
    If HotKeyHandle <> 0 Then
        HotKeyDblClickSupported = CBool((GetClassLong(HotKeyHandle, GCL_STYLE) And CS_DBLCLKS) <> 0)
        If HotKeyBackColorBrush = 0 Then HotKeyBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
        Call ComCtlsSetSubclass(HotKeyHandle, Me, 1)
    End If
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
Else
    If HotKeyHandle <> 0 Then
        If HotKeyBackColorBrush = 0 Then HotKeyBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
        Call ComCtlsSetSubclass(HotKeyHandle, Me, 3)
    End If
End If
End Sub

Private Sub DestroyHotKey()
If HotKeyHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(HotKeyHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow HotKeyHandle, SW_HIDE
SetParent HotKeyHandle, 0
DestroyWindow HotKeyHandle
HotKeyHandle = 0
If HotKeyFontHandle <> 0 Then
    DeleteObject HotKeyFontHandle
    HotKeyFontHandle = 0
End If
If HotKeyBackColorBrush <> 0 Then
    DeleteObject HotKeyBackColorBrush
    HotKeyBackColorBrush = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get Value(Optional ByRef Modifiers As Integer) As VBRUN.KeyCodeConstants
Attribute Value.VB_Description = "Returns/sets the virtual key code and modifier keys that define a hot key combination."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "400"
If HotKeyHandle <> 0 Then
    Dim RetVal As Integer
    RetVal = LoWord(SendMessage(HotKeyHandle, HKM_GETHOTKEY, 0, ByVal 0&))
    Value = LoByte(RetVal)
    Modifiers = HiByte(RetVal)
End If
End Property

Public Property Let Value(Optional ByRef Modifiers As Integer, ByVal NewValue As VBRUN.KeyCodeConstants)
If HotKeyHandle <> 0 Then SendMessage HotKeyHandle, HKM_SETHOTKEY, MakeDWord(MakeWord(NewValue And &HFF&, Modifiers And &HFF&), 0), ByVal 0&
End Property

Public Property Get RawValue() As Long
Attribute RawValue.VB_Description = "Returns/sets the hot key combination."
Attribute RawValue.VB_MemberFlags = "400"
If HotKeyHandle <> 0 Then RawValue = SendMessage(HotKeyHandle, HKM_GETHOTKEY, 0, ByVal 0&)
End Property

Public Property Let RawValue(ByVal Value As Long)
If HotKeyHandle <> 0 Then SendMessage HotKeyHandle, HKM_SETHOTKEY, Value, ByVal 0&
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns the text contained in an object."
Attribute Text.VB_MemberFlags = "400"
If HotKeyHandle <> 0 Then
    Dim hKL As Long
    hKL = GetKeyboardLayout(0)
    Dim RetVal As Integer, KeyCode As Integer, Modifiers As Integer
    RetVal = LoWord(SendMessage(HotKeyHandle, HKM_GETHOTKEY, 0, ByVal 0&))
    KeyCode = LoByte(RetVal)
    Modifiers = HiByte(RetVal)
    Dim ScanCode As Long
    ScanCode = MapVirtualKeyEx(KeyCode, MAPVK_VK_TO_VSC, hKL)
    Dim Buffer As String, StrKey As String
    Buffer = String$(100, vbNullChar)
    GetKeyNameText MakeDWord(0, MakeWord(LoByte(LoWord(ScanCode)), IIf((Modifiers And HOTKEYF_EXT) = HOTKEYF_EXT, 1, 0))), StrPtr(Buffer), 100
    StrKey = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    Dim StrCtrl As String, StrShift As String, StrAlt As String
    Buffer = String$(100, vbNullChar)
    ScanCode = MapVirtualKeyEx(vbKeyControl, MAPVK_VK_TO_VSC, hKL)
    GetKeyNameText MakeDWord(0, MakeWord(LoByte(LoWord(ScanCode)), 0)), StrPtr(Buffer), 100
    StrCtrl = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    Buffer = String$(100, vbNullChar)
    ScanCode = MapVirtualKeyEx(vbKeyShift, MAPVK_VK_TO_VSC, hKL)
    GetKeyNameText MakeDWord(0, MakeWord(LoByte(LoWord(ScanCode)), 0)), StrPtr(Buffer), 100
    StrShift = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    Buffer = String$(100, vbNullChar)
    ScanCode = MapVirtualKeyEx(vbKeyMenu, MAPVK_VK_TO_VSC, hKL)
    GetKeyNameText MakeDWord(0, MakeWord(LoByte(LoWord(ScanCode)), 0)), StrPtr(Buffer), 100
    StrAlt = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    Dim StrModifiers As String
    If (Modifiers And vbCtrlMask) <> 0 Then StrModifiers = StrCtrl
    If (Modifiers And vbShiftMask) <> 0 Then
        If Not StrModifiers = vbNullString Then StrModifiers = StrModifiers & " + " & StrShift Else StrModifiers = StrShift
    End If
    If (Modifiers And vbAltMask) <> 0 Then
        If Not StrModifiers = vbNullString Then StrModifiers = StrModifiers & " + " & StrAlt Else StrModifiers = StrAlt
    End If
    If Not StrModifiers = vbNullString Then Text = StrModifiers & " + " & StrKey Else Text = StrKey
End If
End Property

Public Sub SetRules(ByVal InvalidKeyCombinations As HkeInvalidKeyCombinationConstants, Optional ByVal DefaultModifiers As VBRUN.ShiftConstants)
Attribute SetRules.VB_Description = "Defines the invalid key combinations and the default modifiers for the hot key control."
Select Case InvalidKeyCombinations
    Case HkeInvalidKeyCombinationNone, HkeInvalidKeyCombinationShift, HkeInvalidKeyCombinationCtrl, HkeInvalidKeyCombinationAlt, HkeInvalidKeyCombinationShiftCtrl, HkeInvalidKeyCombinationShiftAlt, HkeInvalidKeyCombinationCtrlAlt, HkeInvalidKeyCombinationShiftCtrlAlt
        If HotKeyHandle <> 0 Then SendMessage HotKeyHandle, HKM_SETRULES, InvalidKeyCombinations, ByVal CLng(DefaultModifiers)
    Case Else
        Err.Raise 380
End Select
End Sub

Public Function SetApplicationHotKey(Optional ByVal hWnd As Long) As Long
Attribute SetApplicationHotKey.VB_Description = "Sets the hot key of the specified window to the current hot key. A hot key cannot be associated with a child window."
If HotKeyHandle <> 0 Then
    If hWnd = 0 Then hWnd = GetAncestor(HotKeyHandle, GA_ROOT)
    SetApplicationHotKey = SendMessage(hWnd, WM_SETHOTKEY, SendMessage(HotKeyHandle, HKM_GETHOTKEY, 0, ByVal 0&), ByVal 0&)
End If
End Function

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
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            HotKeyCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If HotKeyCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(HotKeyCharCodeCache And &HFFFF&)
            HotKeyCharCodeCache = 0
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
    Case WM_ERASEBKGND
        If HotKeyBackColorBrush <> 0 Then
            SetBkMode wParam, 1
            Dim RC As RECT
            GetClientRect hWnd, RC
            FillRect wParam, RC, HotKeyBackColorBrush
            WindowProcControl = 1
            Exit Function
        End If
    
    #If ImplementThemedBorder = True Then
    
    Case WM_THEMECHANGED, WM_STYLECHANGED, WM_ENABLE
        If wMsg = WM_THEMECHANGED Then HotKeyEnabledVisualStyles = EnabledVisualStyles()
        If PropBorderStyle = CCBorderStyleSunken And PropVisualStyles = True Then
            If HotKeyEnabledVisualStyles = True Then SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
        End If
    Case WM_NCPAINT
        ' Bugfix for msctls_hotkey32 class as it always draws a themed border.
        ' However, it should only draw themed when the border style is sunken. (like the edit control)
        ' In addition the disabled and focused state will be handled.
        If PropBorderStyle = CCBorderStyleSunken And PropVisualStyles = True And HotKeyEnabledVisualStyles = True Then
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
                        If HotKeyFocused = True Then
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
                WindowProcControl = 0
                Exit Function
            End If
        End If
        WindowProcControl = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    
    #Else
    
    Case WM_NCPAINT
        WindowProcControl = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    
    #End If
    
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
        If wMsg = WM_LBUTTONDOWN Then
            If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
        End If
        If HotKeyDblClickSupported = False Then
            If HotKeyDblClickTickCount = 0 Then
                HotKeyDblClickTickCount = CLngToULng(GetTickCount())
                HotKeyDblClickX = Get_X_lParam(lParam)
                HotKeyDblClickY = Get_Y_lParam(lParam)
            Else
                If (CLngToULng(GetTickCount()) - HotKeyDblClickTickCount) <= HotKeyDblClickTime Then
                    Dim DblClickRect As RECT
                    With DblClickRect
                    .Left = HotKeyDblClickX - (HotKeyDblClickCX / 2)
                    .Right = HotKeyDblClickX + (HotKeyDblClickCX / 2)
                    .Top = HotKeyDblClickY - (HotKeyDblClickCY / 2)
                    .Bottom = HotKeyDblClickY + (HotKeyDblClickCY / 2)
                    End With
                    If PtInRect(DblClickRect, Get_X_lParam(lParam), Get_Y_lParam(lParam)) <> 0 Then HotKeyIsDblClick = True
                End If
                HotKeyDblClickTickCount = CLngToULng(GetTickCount())
                HotKeyDblClickX = Get_X_lParam(lParam)
                HotKeyDblClickY = Get_Y_lParam(lParam)
            End If
            If HotKeyIsDblClick = True Then
                Select Case wMsg
                    Case WM_LBUTTONDOWN
                        wMsg = WM_LBUTTONDBLCLK
                    Case WM_MBUTTONDOWN
                        wMsg = WM_MBUTTONDBLCLK
                    Case WM_RBUTTONDOWN
                        wMsg = WM_RBUTTONDBLCLK
                End Select
                HotKeyIsDblClick = False
            End If
        End If
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    
    #If ImplementThemedBorder = True Then
    
    Case WM_SETFOCUS, WM_KILLFOCUS
        HotKeyFocused = CBool(wMsg = WM_SETFOCUS)
        If PropBorderStyle = CCBorderStyleSunken And PropVisualStyles = True Then
            If HotKeyEnabledVisualStyles = True Then SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
        End If
    
    #End If
    
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
                HotKeyIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                HotKeyIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                HotKeyIsClick = True
            Case WM_MOUSEMOVE
                If HotKeyMouseOver = False And PropMouseTrack = True Then
                    HotKeyMouseOver = True
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
                If HotKeyIsClick = True Then
                    HotKeyIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If HotKeyMouseOver = True Then
            HotKeyMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        If lParam = HotKeyHandle Then
            Select Case HiWord(wParam)
                Case EN_CHANGE
                    RaiseEvent Change
            End Select
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI HotKeyHandle
End Function

Private Function WindowProcControlDesignMode(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    
    #If ImplementThemedBorder = True Then
    
    Case WM_THEMECHANGED, WM_STYLECHANGED, WM_ENABLE
        WindowProcControlDesignMode = WindowProcControl(hWnd, wMsg, wParam, lParam)
        Exit Function
    
    #End If
    
    Case WM_ERASEBKGND, WM_NCPAINT
        WindowProcControlDesignMode = WindowProcControl(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
WindowProcControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
