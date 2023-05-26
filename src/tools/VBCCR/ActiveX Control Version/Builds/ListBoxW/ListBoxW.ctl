VERSION 5.00
Begin VB.UserControl ListBoxW 
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   PropertyPages   =   "ListBoxW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ListBoxW.ctx":0035
End
Attribute VB_Name = "ListBoxW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const ImplementThemedButton = True

#If False Then
Private LstStyleStandard, LstStyleCheckbox, LstStyleOption
Private LstDrawModeNormal, LstDrawModeOwnerDrawFixed, LstDrawModeOwnerDrawVariable
#End If
Public Enum LstStyleConstants
LstStyleStandard = 0
LstStyleCheckbox = 1
LstStyleOption = 2
End Enum
Public Enum LstDrawModeConstants
LstDrawModeNormal = 0
LstDrawModeOwnerDrawFixed = 1
LstDrawModeOwnerDrawVariable = 2
End Enum
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type TEXTMETRIC
TMHeight As Long
TMAscent As Long
TMDescent As Long
TMInternalLeading As Long
TMExternalLeading As Long
TMAveCharWidth As Long
TMMaxCharWidth As Long
TMWeight As Long
TMOverhang As Long
TMDigitizedAspectX As Long
TMDigitizedAspectY As Long
TMFirstChar As Integer
TMLastChar As Integer
TMDefaultChar As Integer
TMBreakChar As Integer
TMItalic As Byte
TMUnderlined As Byte
TMStruckOut As Byte
TMPitchAndFamily As Byte
TMCharset As Byte
End Type
Private Type MEASUREITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
ItemWidth As Long
ItemHeight As Long
ItemData As Long
End Type
Private Type DRAWITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
ItemAction As Long
ItemState As Long
hWndItem As Long
hDC As Long
RCItem As RECT
ItemData As Long
End Type
Private Type SCROLLINFO
cbSize As Long
fMask As Long
nMin As Long
nMax As Long
nPage As Long
nPos As Long
nTrackPos As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Scroll()
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Public Event ContextMenu(ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event ItemBeforeCheck(ByVal Item As Long, ByRef Cancel As Boolean)
Attribute ItemBeforeCheck.VB_Description = "Occurs when the style property is set to 1 (checkbox) or 2 (option) and before an item is about to be checked or cleared."
Public Event ItemCheck(ByVal Item As Long)
Attribute ItemCheck.VB_Description = "Occurs when the style property is set to 1 (checkbox) or 2 (option) and an item is checked or cleared."
Public Event ItemMeasure(ByVal Item As Long, ByRef ItemHeight As Long)
Attribute ItemMeasure.VB_Description = "Occurs each time an variable owner-drawn list box item's size needs to be determined in preparation of drawing it."
Public Event ItemDraw(ByVal Item As Long, ByVal ItemAction As Long, ByVal ItemState As Long, ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Attribute ItemDraw.VB_Description = "Occurs when a visual aspect of an owner-drawn list box has changed."
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
Private Declare Function LBItemFromPt Lib "comctl32" (ByVal hLB As Long, ByVal PX As Long, ByVal PY As Long, ByVal bAutoScroll As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function DragDetect Lib "user32" (ByVal hWnd As Long, ByVal PX As Integer, ByVal PY As Integer) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (ByRef lpRect As RECT) As Long
Private Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal fMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutW" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutW" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal nTabPositions As Long, ByVal lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal nCtlType As Long, ByVal nFlags As Long) As Long

#If ImplementThemedButton = True Then

Private Enum UxThemeButtonParts
BP_PUSHBUTTON = 1
BP_RADIOBUTTON = 2
BP_CHECKBOX = 3
BP_GROUPBOX = 4
BP_USERBUTTON = 5
End Enum
Private Enum UxThemeCheckBoxStates
CBS_UNCHECKEDNORMAL = 1
CBS_UNCHECKEDHOT = 2
CBS_UNCHECKEDPRESSED = 3
CBS_UNCHECKEDDISABLED = 4
CBS_CHECKEDNORMAL = 5
CBS_CHECKEDHOT = 6
CBS_CHECKEDPRESSED = 7
CBS_CHECKEDDISABLED = 8
End Enum
Private Enum UxThemeRadioButtonStates
RBS_UNCHECKEDNORMAL = 1
RBS_UNCHECKEDHOT = 2
RBS_UNCHECKEDPRESSED = 3
RBS_UNCHECKEDDISABLED = 4
RBS_CHECKEDNORMAL = 5
RBS_CHECKEDHOT = 6
RBS_CHECKEDPRESSED = 7
RBS_CHECKEDDISABLED = 8
End Enum
Private Declare Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme" (ByVal Theme As Long, iPartId As Long, iStateId As Long) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme" (ByVal hWnd As Long, ByVal hDC As Long, ByRef pRect As RECT) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef pClipRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal Theme As Long) As Long

#End If

Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const CF_UNICODETEXT As Long = 13
Private Const TA_RTLREADING = &H100, TA_RIGHT As Long = &H2
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const SW_HIDE As Long = &H0
Private Const WM_MOUSEWHEEL As Long = &H20A
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
Private Const WM_COMMAND As Long = &H111
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_DRAWITEM As Long = &H2B, ODT_LISTBOX As Long = &H2, ODS_SELECTED As Long = &H1, ODS_DISABLED As Long = &H4, ODS_FOCUS As Long = &H10
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_PAINT As Long = &HF
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114
Private Const SB_HORZ As Long = 0
Private Const SB_VERT As Long = 1
Private Const SB_THUMBPOSITION As Long = 4, SB_THUMBTRACK As Long = 5
Private Const SB_LINELEFT As Long = 0, SB_LINERIGHT As Long = 1
Private Const SB_LINEUP As Long = 0, SB_LINEDOWN As Long = 1
Private Const SIF_RANGE As Long = &H1
Private Const SIF_POS As Long = &H4
Private Const SIF_TRACKPOS As Long = &H10
Private Const RGN_COPY As Long = 5
Private Const DFC_BUTTON As Long = &H4, DFCS_BUTTONCHECK As Long = &H0, DFCS_BUTTONRADIO As Long = &H4, DFCS_INACTIVE As Long = &H100, DFCS_CHECKED As Long = &H400, DFCS_FLAT As Long = &H4000
Private Const LB_ERR As Long = (-1)
Private Const LB_ADDSTRING As Long = &H180
Private Const LB_INSERTSTRING As Long = &H181
Private Const LB_DELETESTRING As Long = &H182
Private Const LB_SELITEMRANGEEX As Long = &H183
Private Const LB_RESETCONTENT As Long = &H184
Private Const LB_SETSEL As Long = &H185
Private Const LB_SETCURSEL As Long = &H186
Private Const LB_GETSEL As Long = &H187
Private Const LB_GETCURSEL As Long = &H188
Private Const LB_GETTEXT As Long = &H189
Private Const LB_GETTEXTLEN As Long = &H18A
Private Const LB_GETCOUNT As Long = &H18B
Private Const LB_SELECTSTRING As Long = &H18C
Private Const LB_GETTOPINDEX As Long = &H18E
Private Const LB_FINDSTRING As Long = &H18F
Private Const LB_GETSELCOUNT As Long = &H190
Private Const LB_GETSELITEMS As Long = &H191
Private Const LB_GETHORIZONTALEXTENT As Long = &H193
Private Const LB_SETHORIZONTALEXTENT As Long = &H194
Private Const LB_SETCOLUMNWIDTH As Long = &H195
Private Const LB_SETTOPINDEX As Long = &H197
Private Const LB_GETITEMRECT As Long = &H198
Private Const LB_GETITEMDATA As Long = &H199
Private Const LB_SETITEMDATA As Long = &H19A
Private Const LB_SELITEMRANGE As Long = &H19B ' 16 bit
Private Const LB_SETANCHORINDEX As Long = &H19C
Private Const LB_GETANCHORINDEX As Long = &H19D
Private Const LB_SETCARETINDEX As Long = &H19E
Private Const LB_GETCARETINDEX As Long = &H19F
Private Const LB_SETITEMHEIGHT As Long = &H1A0
Private Const LB_GETITEMHEIGHT As Long = &H1A1
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const LB_ITEMFROMPOINT As Long = &H1A9 ' 16 bit
Private Const LB_GETLISTBOXINFO As Long = &H1B2
Private Const LBS_NOTIFY As Long = &H1
Private Const LBS_SORT As Long = &H2
Private Const LBS_NOREDRAW As Long = &H4
Private Const LBS_MULTIPLESEL As Long = &H8
Private Const LBS_OWNERDRAWFIXED As Long = &H10
Private Const LBS_OWNERDRAWVARIABLE As Long = &H20
Private Const LBS_HASSTRINGS As Long = &H40
Private Const LBS_USETABSTOPS As Long = &H80
Private Const LBS_NOINTEGRALHEIGHT As Long = &H100
Private Const LBS_MULTICOLUMN As Long = &H200
Private Const LBS_EXTENDEDSEL As Long = &H800
Private Const LBS_DISABLENOSCROLL As Long = &H1000
Private Const LBS_NOSEL As Long = &H4000
Private Const LBN_SELCHANGE As Long = 1
Private Const LBN_DBLCLK As Long = 2
Private Const LBN_SELCANCEL As Long = 3
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private ListBoxHandle As Long
Private ListBoxFontHandle As Long
Private ListBoxCharCodeCache As Long
Private ListBoxMouseOver As Boolean
Private ListBoxDesignMode As Boolean
Private ListBoxNewIndex As Long
Private ListBoxDragIndexBuffer As Long, ListBoxDragIndex As Long
Private ListBoxTopIndex As Long
Private ListBoxInsertMark As Long, ListBoxInsertMarkAfter As Boolean
Private ListBoxItemCheckedCount As Long
Private ListBoxItemChecked() As Byte, ListBoxOptionIndex As Long
Private ListBoxStateImageSize As Long
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropOLEDragDropScroll As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropRedraw As Boolean
Private PropBorderStyle As CCBorderStyleConstants
Private PropMultiColumn As Boolean
Private PropSorted As Boolean
Private PropIntegralHeight As Boolean
Private PropAllowSelection As Boolean
Private PropMultiSelect As VBRUN.MultiSelectConstants
Private PropHorizontalExtent As Long
Private PropUseTabStops As Boolean
Private PropStyle As LstStyleConstants
Private PropDisableNoScroll As Boolean
Private PropDrawMode As LstDrawModeConstants
Private PropInsertMarkColor As OLE_COLOR
Private PropScrollTrack As Boolean

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
Call ComCtlsInitCC(ICC_STANDARD_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim ListBoxItemChecked(0) As Byte
ListBoxStateImageSize = (15 * PixelsPerDIP_X())
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
ListBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropOLEDragMode = vbOLEDragManual
PropOLEDragDropScroll = True
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropRedraw = True
PropBorderStyle = CCBorderStyleSunken
PropSorted = False
PropIntegralHeight = True
PropAllowSelection = True
PropMultiSelect = vbMultiSelectNone
PropHorizontalExtent = 0
PropUseTabStops = True
PropStyle = LstStyleStandard
PropDisableNoScroll = False
PropDrawMode = LstDrawModeNormal
PropInsertMarkColor = vbBlack
PropScrollTrack = True
Call CreateListBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
ListBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
Me.ForeColor = .ReadProperty("ForeColor", vbButtonText)
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
PropRedraw = .ReadProperty("Redraw", True)
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSunken)
PropMultiColumn = .ReadProperty("MultiColumn", False)
PropSorted = .ReadProperty("Sorted", False)
PropIntegralHeight = .ReadProperty("IntegralHeight", True)
PropAllowSelection = .ReadProperty("AllowSelection", True)
PropMultiSelect = .ReadProperty("MultiSelect", vbMultiSelectNone)
PropHorizontalExtent = .ReadProperty("HorizontalExtent", 0)
PropUseTabStops = .ReadProperty("UseTabStops", True)
PropStyle = .ReadProperty("Style", LstStyleStandard)
PropDisableNoScroll = .ReadProperty("DisableNoScroll", False)
PropDrawMode = .ReadProperty("DrawMode", LstDrawModeNormal)
PropInsertMarkColor = .ReadProperty("InsertMarkColor", vbBlack)
PropScrollTrack = .ReadProperty("ScrollTrack", True)
End With
Call CreateListBox
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "ForeColor", Me.ForeColor, vbButtonText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDragDropScroll", PropOLEDragDropScroll, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "Redraw", PropRedraw, True
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSunken
.WriteProperty "MultiColumn", PropMultiColumn, False
.WriteProperty "Sorted", PropSorted, False
.WriteProperty "IntegralHeight", PropIntegralHeight, True
.WriteProperty "AllowSelection", PropAllowSelection, True
.WriteProperty "MultiSelect", PropMultiSelect, vbMultiSelectNone
.WriteProperty "HorizontalExtent", PropHorizontalExtent, 0
.WriteProperty "UseTabStops", PropUseTabStops, True
.WriteProperty "Style", PropStyle, LstStyleStandard
.WriteProperty "DisableNoScroll", PropDisableNoScroll, False
.WriteProperty "DrawMode", PropDrawMode, LstDrawModeNormal
.WriteProperty "InsertMarkColor", PropInsertMarkColor, vbBlack
.WriteProperty "ScrollTrack", PropScrollTrack, True
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
ListBoxDragIndex = 0
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
If ListBoxHandle <> 0 Then
    If State = vbOver And Not Effect = vbDropEffectNone Then
        If PropOLEDragDropScroll = True Then
            Dim RC As RECT
            GetWindowRect ListBoxHandle, RC
            Dim dwStyle As Long
            dwStyle = GetWindowLong(ListBoxHandle, GWL_STYLE)
            If (dwStyle And WS_HSCROLL) = WS_HSCROLL Then
                If Abs(X) < (16 * PixelsPerDIP_X()) Then
                    SendMessage ListBoxHandle, WM_HSCROLL, SB_LINELEFT, ByVal 0&
                ElseIf Abs(X - (RC.Right - RC.Left)) < (16 * PixelsPerDIP_X()) Then
                    SendMessage ListBoxHandle, WM_HSCROLL, SB_LINERIGHT, ByVal 0&
                End If
            End If
            If (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                If Abs(Y) < (16 * PixelsPerDIP_Y()) Then
                    SendMessage ListBoxHandle, WM_VSCROLL, SB_LINEUP, ByVal 0&
                ElseIf Abs(Y - (RC.Bottom - RC.Top)) < (16 * PixelsPerDIP_Y()) Then
                    SendMessage ListBoxHandle, WM_VSCROLL, SB_LINEDOWN, ByVal 0&
                End If
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
If ListBoxDragIndex > 0 Then
    If PropOLEDragMode = vbOLEDragAutomatic Then
        Dim SelIndices As Collection, Text As String
        Set SelIndices = Me.SelectedIndices
        With SelIndices
        If .Count > 0 Then
            Dim Item As Variant, i As Long
            For Each Item In SelIndices
                i = i + 1
                Text = Text & Me.List(Item) & IIf(i < .Count, vbCrLf, vbNullString)
            Next Item
        End If
        End With
        Data.SetData StrToVar(Text & vbNullChar), CF_UNICODETEXT
        Data.SetData Text, vbCFText
        AllowedEffects = vbDropEffectCopy
    End If
ElseIf ListBoxHandle <> 0 Then
    Dim P As POINTAPI
    GetCursorPos P
    ListBoxDragIndex = LBItemFromPt(ListBoxHandle, P.X, P.Y, 0) + 1
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then ListBoxDragIndex = 0
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
If ListBoxDragIndex > 0 Then Exit Sub
If ListBoxDragIndexBuffer > 0 Then ListBoxDragIndex = ListBoxDragIndexBuffer
UserControl.OLEDrag
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If ListBoxDesignMode = True And PropertyName = "DisplayName" Then
    If ListBoxHandle <> 0 Then
        If SendMessage(ListBoxHandle, LB_GETCOUNT, 0, ByVal 0&) > 0 Then
            Dim Buffer As String
            Buffer = Ambient.DisplayName
            SendMessage ListBoxHandle, LB_RESETCONTENT, 0, ByVal 0&
            SendMessage ListBoxHandle, LB_ADDSTRING, 0, ByVal StrPtr(Buffer)
            SendMessage ListBoxHandle, LB_SETCURSEL, -1, ByVal 0&
        End If
    End If
End If
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If ListBoxHandle = 0 Then InProc = False: Exit Sub
Dim WndRect As RECT
MoveWindow ListBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
If PropIntegralHeight = True Then
    GetWindowRect ListBoxHandle, WndRect
    .Extender.Height = .ScaleY((WndRect.Bottom - WndRect.Top), vbPixels, vbContainerSize)
End If
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyListBox
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
hWnd = ListBoxHandle
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
OldFontHandle = ListBoxFontHandle
ListBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ListBoxHandle <> 0 Then SendMessage ListBoxHandle, WM_SETFONT, ListBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If PropStyle <> LstStyleStandard And ListBoxHandle <> 0 Then
    Dim hDCScreen As Long
    hDCScreen = GetDC(0)
    If hDCScreen <> 0 Then
        Dim TM As TEXTMETRIC, hFontOld As Long
        If ListBoxFontHandle <> 0 Then hFontOld = SelectObject(hDCScreen, ListBoxFontHandle)
        If GetTextMetrics(hDCScreen, TM) <> 0 Then
            If TM.TMHeight < ListBoxStateImageSize Then TM.TMHeight = ListBoxStateImageSize
            SendMessage ListBoxHandle, LB_SETITEMHEIGHT, 0, ByVal TM.TMHeight
            If PropIntegralHeight = True Then
                MoveWindow ListBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + 1, 0
                MoveWindow ListBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 0
            End If
        End If
        If hFontOld <> 0 Then SelectObject hDCScreen, hFontOld
        ReleaseDC 0, hDCScreen
    End If
End If
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = ListBoxFontHandle
ListBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ListBoxHandle <> 0 Then SendMessage ListBoxHandle, WM_SETFONT, ListBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If PropStyle <> LstStyleStandard And ListBoxHandle <> 0 Then
    Dim hDCScreen As Long
    hDCScreen = GetDC(0)
    If hDCScreen <> 0 Then
        Dim TM As TEXTMETRIC, hFontOld As Long
        If ListBoxFontHandle <> 0 Then hFontOld = SelectObject(hDCScreen, ListBoxFontHandle)
        If GetTextMetrics(hDCScreen, TM) <> 0 Then
            If TM.TMHeight < ListBoxStateImageSize Then TM.TMHeight = ListBoxStateImageSize
            SendMessage ListBoxHandle, LB_SETITEMHEIGHT, 0, ByVal TM.TMHeight
            If PropIntegralHeight = True Then
                MoveWindow ListBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + 1, 0
                MoveWindow ListBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 0
            End If
        End If
        If hFontOld <> 0 Then SelectObject hDCScreen, hFontOld
        ReleaseDC 0, hDCScreen
    End If
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
If ListBoxHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles ListBoxHandle
    Else
        RemoveVisualStyles ListBoxHandle
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
If ListBoxHandle <> 0 Then EnableWindow ListBoxHandle, IIf(Value = True, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDragMode() As VBRUN.OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this control can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
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
If ListBoxDesignMode = False Then Call RefreshMousePointer
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
        If ListBoxDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ListBoxDesignMode = False Then Call RefreshMousePointer
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
If PropRightToLeft = True Then dwMask = WS_EX_RTLREADING Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR
If ListBoxHandle <> 0 Then Call ComCtlsSetRightToLeft(ListBoxHandle, dwMask)
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

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Returns/sets a value that determines whether or not the list box redraws when changing the items. You can speed up the creation of large lists by disabling this property before adding the items."
Redraw = PropRedraw
End Property

Public Property Let Redraw(ByVal Value As Boolean)
PropRedraw = Value
If ListBoxHandle <> 0 And ListBoxDesignMode = False Then
    SendMessage ListBoxHandle, WM_SETREDRAW, IIf(PropRedraw = True, 1, 0), ByVal 0&
    If PropRedraw = True Then Me.Refresh
End If
UserControl.PropertyChanged "Redraw"
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
If ListBoxHandle <> 0 Then
    Call ComCtlsChangeBorderStyle(ListBoxHandle, PropBorderStyle)
    Call UserControl_Resize
End If
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get MultiColumn() As Boolean
Attribute MultiColumn.VB_Description = "Returns/sets a value that determines whether or not the control is scrolled horizontally and the items are displayed in multiple columns."
MultiColumn = PropMultiColumn
End Property

Public Property Let MultiColumn(ByVal Value As Boolean)
If PropDrawMode = LstDrawModeOwnerDrawVariable And Value = True Then
    If ListBoxDesignMode = True Then
        MsgBox "MultiColumn must be False when DrawMode is 2 - OwnerDrawVariable", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=383, Description:="MultiColumn must be False when DrawMode is 2 - OwnerDrawVariable"
    End If
End If
PropMultiColumn = Value
If ListBoxHandle <> 0 Then Call ReCreateListBox
UserControl.PropertyChanged "MultiColumn"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
Sorted = PropSorted
End Property

Public Property Let Sorted(ByVal Value As Boolean)
PropSorted = Value
If ListBoxHandle <> 0 Then Call ReCreateListBox
UserControl.PropertyChanged "Sorted"
End Property

Public Property Get IntegralHeight() As Boolean
Attribute IntegralHeight.VB_Description = "Returns/sets a value indicating whether the control displays partial items. This flag is always set to false in an variable owner-drawn list box."
IntegralHeight = PropIntegralHeight
End Property

Public Property Let IntegralHeight(ByVal Value As Boolean)
If ListBoxDesignMode = False Then
    Err.Raise Number:=382, Description:="IntegralHeight property is read-only at run time"
Else
    PropIntegralHeight = Value
    If ListBoxHandle <> 0 Then Call ReCreateListBox
End If
UserControl.PropertyChanged "IntegralHeight"
End Property

Public Property Get AllowSelection() As Boolean
Attribute AllowSelection.VB_Description = "Returns/sets a value indicating if the list box enables selection of items."
AllowSelection = PropAllowSelection
End Property

Public Property Let AllowSelection(ByVal Value As Boolean)
PropAllowSelection = Value
If ListBoxHandle <> 0 Then Call ReCreateListBox
UserControl.PropertyChanged "AllowSelection"
End Property

Public Property Get MultiSelect() As VBRUN.MultiSelectConstants
Attribute MultiSelect.VB_Description = "Returns/sets a value that determines whether a user can make multiple selections in a control."
MultiSelect = PropMultiSelect
End Property

Public Property Let MultiSelect(ByVal Value As VBRUN.MultiSelectConstants)
Select Case Value
    Case vbMultiSelectNone, vbMultiSelectSimple, vbMultiSelectExtended
        If PropStyle <> LstStyleStandard And Value <> vbMultiSelectNone Then
            If ListBoxDesignMode = True Then
                MsgBox "MultiSelect must be 0 - None when Style is not 0 - Standard", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise Number:=383, Description:="MultiSelect must be 0 - None when Style is not 0 - Standard"
            End If
        End If
        PropMultiSelect = Value
    Case Else
        Err.Raise 380
End Select
If ListBoxHandle <> 0 Then Call ReCreateListBox
UserControl.PropertyChanged "MultiSelect"
End Property

Public Property Get HorizontalExtent() As Single
Attribute HorizontalExtent.VB_Description = "Returns/sets the width by which a list box can be scrolled horizontally. This is only meaningful if the multi column property is set to false."
If ListBoxHandle <> 0 And PropMultiColumn = False Then
    HorizontalExtent = UserControl.ScaleX(SendMessage(ListBoxHandle, LB_GETHORIZONTALEXTENT, 0, ByVal 0&), vbPixels, vbContainerSize)
Else
    HorizontalExtent = UserControl.ScaleX(PropHorizontalExtent, vbPixels, vbContainerSize)
End If
End Property

Public Property Let HorizontalExtent(ByVal Value As Single)
If Value < 0 Then
    If ListBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropHorizontalExtent = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If ListBoxHandle <> 0 And PropMultiColumn = False Then SendMessage ListBoxHandle, LB_SETHORIZONTALEXTENT, PropHorizontalExtent, ByVal 0&
UserControl.PropertyChanged "HorizontalExtent"
End Property

Public Property Get UseTabStops() As Boolean
Attribute UseTabStops.VB_Description = "Returns/sets a value indicating if the list box can recognize and expand tab characters when drawing its strings."
UseTabStops = PropUseTabStops
End Property

Public Property Let UseTabStops(ByVal Value As Boolean)
PropUseTabStops = Value
If ListBoxHandle <> 0 Then Call ReCreateListBox
UserControl.PropertyChanged "UseTabStops"
End Property

Public Property Get Style() As LstStyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines whether checkboxes or options are displayed."
Style = PropStyle
End Property

Public Property Let Style(ByVal Value As LstStyleConstants)
If ListBoxDesignMode = False Then
    Err.Raise Number:=382, Description:="Style property is read-only at run time"
Else
    Select Case Value
        Case LstStyleStandard, LstStyleCheckbox, LstStyleOption
            If PropDrawMode <> LstDrawModeNormal And Value <> LstStyleStandard Then
                MsgBox "Style must be 0 - Standard when DrawMode is not 0 - Normal", vbCritical + vbOKOnly
                Exit Property
            End If
            PropStyle = Value
            If PropStyle <> LstStyleStandard Then PropMultiSelect = vbMultiSelectNone
        Case Else
            Err.Raise 380
    End Select
    If ListBoxHandle <> 0 Then Call ReCreateListBox
End If
UserControl.PropertyChanged "Style"
End Property

Public Property Get DisableNoScroll() As Boolean
Attribute DisableNoScroll.VB_Description = "Returns/sets a value that determines whether scroll bars are disabled instead of hided when they are not needed."
DisableNoScroll = PropDisableNoScroll
End Property

Public Property Let DisableNoScroll(ByVal Value As Boolean)
PropDisableNoScroll = Value
If ListBoxHandle <> 0 Then Call ReCreateListBox
UserControl.PropertyChanged "DisableNoScroll"
End Property

Public Property Get DrawMode() As LstDrawModeConstants
Attribute DrawMode.VB_Description = "Returns/sets a value indicating whether your code or the operating system will handle drawing of the elements."
DrawMode = PropDrawMode
End Property

Public Property Let DrawMode(ByVal Value As LstDrawModeConstants)
Select Case Value
    Case LstDrawModeNormal, LstDrawModeOwnerDrawFixed, LstDrawModeOwnerDrawVariable
        If ListBoxDesignMode = False Then
            Err.Raise Number:=382, Description:="DrawMode property is read-only at run time"
        Else
            PropDrawMode = Value
            If PropDrawMode <> LstDrawModeNormal Then PropStyle = LstStyleStandard
            If ListBoxHandle <> 0 Then Call ReCreateListBox
        End If
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "DrawMode"
End Property

Public Property Get InsertMarkColor() As OLE_COLOR
Attribute InsertMarkColor.VB_Description = "Returns/sets the color of the insertion mark."
InsertMarkColor = PropInsertMarkColor
End Property

Public Property Let InsertMarkColor(ByVal Value As OLE_COLOR)
PropInsertMarkColor = Value
If ListBoxInsertMark > -1 Then Call InvalidateInsertMark
UserControl.PropertyChanged "InsertMarkColor"
End Property

Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_Description = "Returns/sets whether the control should scroll its contents while the user moves the scroll box along the scroll bars."
ScrollTrack = PropScrollTrack
End Property

Public Property Let ScrollTrack(ByVal Value As Boolean)
PropScrollTrack = Value
UserControl.PropertyChanged "ScrollTrack"
End Property

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to the list box."
If ListBoxHandle <> 0 Then
    Dim RetVal As Long
    If IsMissing(Index) = True Then
        RetVal = SendMessage(ListBoxHandle, LB_ADDSTRING, 0, ByVal StrPtr(Item))
    Else
        Dim IndexLong As Long
        Select Case VarType(Index)
            Case vbLong, vbInteger, vbByte
                If Index >= 0 Then
                    IndexLong = Index
                Else
                    Err.Raise 5
                End If
            Case vbDouble, vbSingle
                If CLng(Index) >= 0 Then
                    IndexLong = CLng(Index)
                Else
                    Err.Raise 5
                End If
            Case vbString
                IndexLong = CLng(Index)
                If IndexLong < 0 Then Err.Raise 5
            Case Else
                Err.Raise 13
        End Select
        RetVal = SendMessage(ListBoxHandle, LB_INSERTSTRING, IndexLong, ByVal StrPtr(Item))
    End If
    If Not RetVal = LB_ERR Then
        ListBoxNewIndex = RetVal
        If PropStyle <> LstStyleStandard Then
            ListBoxItemCheckedCount = ListBoxItemCheckedCount + 1
            If PropStyle = LstStyleCheckbox Then
                ReDim Preserve ListBoxItemChecked(0 To ListBoxItemCheckedCount) As Byte
                If ListBoxNewIndex < (ListBoxItemCheckedCount - 1) Then CopyMemory ByVal VarPtr(ListBoxItemChecked(ListBoxNewIndex + 2)), ByVal VarPtr(ListBoxItemChecked(ListBoxNewIndex + 1)), (ListBoxItemCheckedCount - ListBoxNewIndex - 1)
                ListBoxItemChecked(ListBoxNewIndex + 1) = vbUnchecked
            ElseIf PropStyle = LstStyleOption Then
                If ListBoxNewIndex <= ListBoxOptionIndex Then ListBoxOptionIndex = ListBoxOptionIndex + 1
            End If
        End If
    Else
        Err.Raise 5
    End If
End If
End Sub

Public Sub RemoveItem(ByVal Index As Long)
Attribute RemoveItem.VB_Description = "Removes an item from the list box."
If ListBoxHandle <> 0 Then
    If Index >= 0 Then
        If Not SendMessage(ListBoxHandle, LB_DELETESTRING, Index, ByVal 0&) = LB_ERR Then
            ListBoxNewIndex = -1
            If ListBoxInsertMark > -1 Then
                If ListBoxInsertMark > (SendMessage(ListBoxHandle, LB_GETCOUNT, 0, ByVal 0&) - 1) Then
                    ListBoxInsertMark = -1
                    ListBoxInsertMarkAfter = False
                End If
            End If
            If PropStyle <> LstStyleStandard Then
                ListBoxItemCheckedCount = ListBoxItemCheckedCount - 1
                If PropStyle = LstStyleCheckbox Then
                    If ListBoxItemCheckedCount > 0 Then
                        If Index < ListBoxItemCheckedCount Then CopyMemory ByVal VarPtr(ListBoxItemChecked(Index + 1)), ByVal VarPtr(ListBoxItemChecked(Index + 2)), (ListBoxItemCheckedCount - Index)
                        ReDim Preserve ListBoxItemChecked(0 To ListBoxItemCheckedCount) As Byte
                    Else
                        ReDim ListBoxItemChecked(0) As Byte
                    End If
                ElseIf PropStyle = LstStyleOption Then
                    If ListBoxOptionIndex > -1 Then
                        If ListBoxItemCheckedCount > 0 Then
                            If ListBoxOptionIndex > (SendMessage(ListBoxHandle, LB_GETCOUNT, 0, ByVal 0&) - 1) Then
                                ListBoxOptionIndex = -1
                            ElseIf Index = ListBoxOptionIndex Then
                                ListBoxOptionIndex = -1
                            ElseIf Index < ListBoxOptionIndex Then
                                ListBoxOptionIndex = ListBoxOptionIndex - 1
                            End If
                        Else
                            ListBoxOptionIndex = -1
                        End If
                    End If
                End If
            End If
        Else
            Err.Raise 5
        End If
    Else
        Err.Raise 5
    End If
End If
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of the list box."
If ListBoxHandle <> 0 Then
    SendMessage ListBoxHandle, LB_RESETCONTENT, 0, ByVal 0&
    ListBoxNewIndex = -1
    If PropStyle <> LstStyleStandard Then
        ListBoxItemCheckedCount = 0
        ReDim ListBoxItemChecked(0) As Byte
        ListBoxOptionIndex = -1
    End If
End If
End Sub

Public Property Get ListCount() As Long
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
If ListBoxHandle <> 0 Then ListCount = SendMessage(ListBoxHandle, LB_GETCOUNT, 0, ByVal 0&)
End Property

Public Property Get List(ByVal Index As Long) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
Attribute List.VB_MemberFlags = "400"
If ListBoxHandle <> 0 Then
    Dim Length As Long
    Length = SendMessage(ListBoxHandle, LB_GETTEXTLEN, Index, ByVal 0&)
    If Not Length = LB_ERR Then
        List = String(Length, vbNullChar)
        SendMessage ListBoxHandle, LB_GETTEXT, Index, ByVal StrPtr(List)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Let List(ByVal Index As Long, ByVal Value As String)
If ListBoxHandle <> 0 Then
    If Index > -1 Then
        Dim ListIndex As Long, SelVal As Long, ItemData As Long
        ListIndex = Me.ListIndex
        If PropMultiSelect <> vbMultiSelectNone Then SelVal = SendMessage(ListBoxHandle, LB_GETSEL, Index, ByVal 0&)
        ItemData = SendMessage(ListBoxHandle, LB_GETITEMDATA, Index, ByVal 0&)
        If Not SendMessage(ListBoxHandle, LB_DELETESTRING, Index, ByVal 0&) = LB_ERR Then
            SendMessage ListBoxHandle, LB_INSERTSTRING, Index, ByVal StrPtr(Value)
            Me.ListIndex = ListIndex
            If PropMultiSelect <> vbMultiSelectNone And Not SelVal = LB_ERR Then SendMessage ListBoxHandle, LB_SETSEL, SelVal, ByVal Index
            SendMessage ListBoxHandle, LB_SETITEMDATA, Index, ByVal ItemData
            On Error Resume Next
            UserControl.Extender.DataChanged = True
            On Error GoTo 0
        Else
            Err.Raise 5
        End If
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
Attribute ListIndex.VB_MemberFlags = "400"
If ListBoxHandle <> 0 Then
    If PropMultiSelect = vbMultiSelectNone Then
        ListIndex = SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&)
    Else
        ListIndex = SendMessage(ListBoxHandle, LB_GETCARETINDEX, 0, ByVal 0&)
    End If
End If
End Property

Public Property Let ListIndex(ByVal Value As Long)
If ListBoxHandle <> 0 Then
    Dim Changed As Boolean
    If PropMultiSelect = vbMultiSelectNone Then
        Changed = CBool(SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&) <> Value)
        If Not Value = -1 Then
            If SendMessage(ListBoxHandle, LB_SETCURSEL, Value, ByVal 0&) = LB_ERR Then Err.Raise 380
        Else
            SendMessage ListBoxHandle, LB_SETCURSEL, -1, ByVal 0&
        End If
    Else
        Changed = CBool(SendMessage(ListBoxHandle, LB_GETCARETINDEX, 0, ByVal 0&) <> Value)
        If SendMessage(ListBoxHandle, LB_SETCARETINDEX, Value, ByVal 0&) = LB_ERR Then Err.Raise 380
    End If
    If Changed = True Then RaiseEvent Click
End If
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a list box."
Attribute ItemData.VB_MemberFlags = "400"
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Index, ByVal 0&) = LB_ERR Then
        ItemData = SendMessage(ListBoxHandle, LB_GETITEMDATA, Index, ByVal 0&)
    Else
        Err.Raise 381
    End If
End If
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal Value As Long)
If ListBoxHandle <> 0 Then If SendMessage(ListBoxHandle, LB_SETITEMDATA, Index, ByVal Value) = LB_ERR Then Err.Raise 381
End Property

Public Property Get ItemChecked(ByVal Index As Long) As Boolean
Attribute ItemChecked.VB_Description = "Returns/sets a value that determines whether the item is checked or not. This is only meaningful if the style property is set to 1 (checkbox) or 2 (option)."
Attribute ItemChecked.VB_MemberFlags = "400"
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Index, ByVal 0&) = LB_ERR Then
        If Index <= (ListBoxItemCheckedCount - 1) Then
            If PropStyle = LstStyleCheckbox Then
                ItemChecked = CBool(ListBoxItemChecked(Index + 1) = vbChecked)
            ElseIf PropStyle = LstStyleOption Then
                ItemChecked = CBool(ListBoxOptionIndex = Index)
            End If
        End If
    Else
        Err.Raise 381
    End If
End If
End Property

Public Property Let ItemChecked(ByVal Index As Long, ByVal Value As Boolean)
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Index, ByVal 0&) = LB_ERR Then
        If Index <= (ListBoxItemCheckedCount - 1) Then
            Dim Changed As Boolean
            If PropStyle = LstStyleCheckbox Then
                Changed = CBool(ListBoxItemChecked(Index + 1) <> IIf(Value = True, vbChecked, vbUnchecked))
            ElseIf PropStyle = LstStyleOption Then
                If ListBoxOptionIndex <> Index Then
                    Changed = Value
                ElseIf Value = False Then
                    Changed = True
                End If
            End If
            If Changed = True Then
                Dim Cancel As Boolean
                RaiseEvent ItemBeforeCheck(Index, Cancel)
                If Cancel = False Then
                    Dim RC As RECT
                    If PropStyle = LstStyleCheckbox Then
                        ListBoxItemChecked(Index + 1) = IIf(Value = True, vbChecked, vbUnchecked)
                    ElseIf PropStyle = LstStyleOption Then
                        If ListBoxOptionIndex > -1 Then
                            SendMessage ListBoxHandle, LB_GETITEMRECT, ListBoxOptionIndex, ByVal VarPtr(RC)
                            InvalidateRect ListBoxHandle, RC, 0
                        End If
                        If ListBoxOptionIndex <> Index Then
                            ListBoxOptionIndex = Index
                        ElseIf Value = False Then
                            ListBoxOptionIndex = -1
                        End If
                    End If
                    SendMessage ListBoxHandle, LB_GETITEMRECT, Index, ByVal VarPtr(RC)
                    InvalidateRect ListBoxHandle, RC, 0
                    RaiseEvent ItemCheck(Index)
                End If
            End If
        End If
    Else
        Err.Raise 381
    End If
End If
End Property

Private Sub CreateListBox()
If ListBoxHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or LBS_NOTIFY Or WS_HSCROLL
If PropRedraw = False Then dwStyle = dwStyle Or LBS_NOREDRAW
Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, PropBorderStyle)
If PropDrawMode = LstDrawModeOwnerDrawVariable Then
    ' The LBS_MULTICOLUMN and LBS_OWNERDRAWVARIABLE styles cannot be combined.
    PropMultiColumn = False
    ' In an variable owner-drawn list box it makes no sense to have an integral height.
    ' Otherwise it would come to unpredictable adjustments.
    PropIntegralHeight = False
End If
If PropMultiColumn = False Then
    dwStyle = dwStyle Or WS_VSCROLL
    If PropRightToLeft = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR
Else
    dwStyle = dwStyle Or LBS_MULTICOLUMN
End If
If PropSorted = True Then dwStyle = dwStyle Or LBS_SORT
If PropIntegralHeight = False Then dwStyle = dwStyle Or LBS_NOINTEGRALHEIGHT
If PropAllowSelection = False Then dwStyle = dwStyle Or LBS_NOSEL
Select Case PropMultiSelect
    Case vbMultiSelectSimple
        dwStyle = dwStyle Or LBS_MULTIPLESEL
    Case vbMultiSelectExtended
        dwStyle = dwStyle Or LBS_EXTENDEDSEL
End Select
If PropUseTabStops = True Then dwStyle = dwStyle Or LBS_USETABSTOPS
If PropStyle <> LstStyleStandard Then dwStyle = dwStyle Or LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS
If PropDisableNoScroll = True Then dwStyle = dwStyle Or LBS_DISABLENOSCROLL
If PropStyle = LstStyleStandard Then
    Select Case PropDrawMode
        Case LstDrawModeOwnerDrawFixed
            dwStyle = dwStyle Or LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS
        Case LstDrawModeOwnerDrawVariable
            dwStyle = dwStyle Or LBS_OWNERDRAWVARIABLE Or LBS_HASSTRINGS
    End Select
End If
ListBoxHandle = CreateWindowEx(dwExStyle, StrPtr("ListBox"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If ListBoxHandle <> 0 Then
    Call ComCtlsShowAllUIStates(ListBoxHandle)
    If PropMultiColumn = True And PropRightToLeft = True Then
        ' In a multi-column list box it is necessary to set the right-to-left alignment afterwards.
        ' Else the top index gets negative and everything will be unpredictable and unstable. (Bug?)
        Call ComCtlsSetRightToLeft(ListBoxHandle, WS_EX_RTLREADING Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR)
    End If
    If PropMultiColumn = False And PropHorizontalExtent > 0 Then SendMessage ListBoxHandle, LB_SETHORIZONTALEXTENT, PropHorizontalExtent, ByVal 0&
    ListBoxNewIndex = -1
    ListBoxTopIndex = 0
    ListBoxInsertMark = -1
    ListBoxInsertMarkAfter = False
    ListBoxOptionIndex = -1
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If ListBoxDesignMode = False Then
    If ListBoxHandle <> 0 Then Call ComCtlsSetSubclass(ListBoxHandle, Me, 1)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
Else
    If ListBoxHandle <> 0 Then
        Dim Buffer As String
        Buffer = Ambient.DisplayName
        SendMessage ListBoxHandle, LB_ADDSTRING, 0, ByVal StrPtr(Buffer)
        SendMessage ListBoxHandle, LB_SETCURSEL, -1, ByVal 0&
    End If
    If PropStyle <> LstStyleStandard Then
        Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
        Me.Refresh
    End If
End If
End Sub

Private Sub ReCreateListBox()
If ListBoxDesignMode = False Then
    Dim Locked As Boolean
    With Me
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Dim ListArr() As String, ItemDataArr() As Long, ItemSelArr() As Long
    Dim ItemHeight As Long, ListIndex As Long, TopIndex As Long, NewIndex As Long, InsertMark As Long, InsertMarkAfter As Boolean
    Dim Count As Long, i As Long
    If ListBoxHandle <> 0 Then
        ItemHeight = SendMessage(ListBoxHandle, LB_GETITEMHEIGHT, 0, ByVal 0&)
        Count = SendMessage(ListBoxHandle, LB_GETCOUNT, 0, ByVal 0&)
        If Count > 0 Then
            ReDim ListArr(0 To (Count - 1)) As String
            ReDim ItemDataArr(0 To (Count - 1)) As Long
            ReDim ItemSelArr(0 To (Count - 1)) As Long
            For i = 0 To (Count - 1)
                ListArr(i) = .List(i)
                ItemDataArr(i) = SendMessage(ListBoxHandle, LB_GETITEMDATA, i, ByVal 0&)
                If PropMultiSelect <> vbMultiSelectNone Then ItemSelArr(i) = SendMessage(ListBoxHandle, LB_GETSEL, i, ByVal 0&)
            Next i
        End If
        ListIndex = .ListIndex
        TopIndex = .TopIndex
    End If
    NewIndex = ListBoxNewIndex
    InsertMark = ListBoxInsertMark
    InsertMarkAfter = ListBoxInsertMarkAfter
    Call DestroyListBox
    Call CreateListBox
    Call UserControl_Resize
    If ListBoxHandle <> 0 Then
        SendMessage ListBoxHandle, LB_SETITEMHEIGHT, 0, ByVal ItemHeight
        If Count > 0 Then
            SendMessage ListBoxHandle, WM_SETREDRAW, 0, ByVal 0&
            For i = 0 To (Count - 1)
                SendMessage ListBoxHandle, LB_INSERTSTRING, i, ByVal StrPtr(ListArr(i))
                SendMessage ListBoxHandle, LB_SETITEMDATA, i, ByVal ItemDataArr(i)
                If PropMultiSelect <> vbMultiSelectNone Then SendMessage ListBoxHandle, LB_SETSEL, ItemSelArr(i), ByVal i
            Next i
            SendMessage ListBoxHandle, WM_SETREDRAW, 1, ByVal 0&
        End If
        .ListIndex = ListIndex
        .TopIndex = TopIndex
    End If
    ListBoxNewIndex = NewIndex
    ListBoxInsertMark = InsertMark
    ListBoxInsertMarkAfter = InsertMarkAfter
    If Locked = True Then LockWindowUpdate 0
    .Refresh
    If PropRedraw = False Then .Redraw = PropRedraw
    End With
Else
    Call DestroyListBox
    Call ComCtlsRemoveSubclass(UserControl.hWnd)
    Call CreateListBox
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyListBox()
If ListBoxHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(ListBoxHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow ListBoxHandle, SW_HIDE
SetParent ListBoxHandle, 0
DestroyWindow ListBoxHandle
ListBoxHandle = 0
If ListBoxFontHandle <> 0 Then
    DeleteObject ListBoxFontHandle
    ListBoxFontHandle = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
If PropRedraw = True Or ListBoxDesignMode = True Then RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "143c"
If ListBoxHandle <> 0 Then
    Dim Index As Long
    Index = Me.ListIndex
    If Index > -1 Then Text = Me.List(Index)
End If
End Property

Public Property Let Text(ByVal Value As String)
If ListBoxHandle <> 0 Then Me.ListIndex = SendMessage(ListBoxHandle, LB_FINDSTRINGEXACT, -1, ByVal StrPtr(Value))
End Property

Public Property Get SelCount() As Long
Attribute SelCount.VB_Description = "Returns the number of selected items in the list box."
Attribute SelCount.VB_MemberFlags = "400"
If ListBoxHandle <> 0 Then
    Dim RetVal As Long
    RetVal = SendMessage(ListBoxHandle, LB_GETSELCOUNT, 0, ByVal 0&)
    If Not RetVal = LB_ERR Then
        SelCount = RetVal
    Else
        RetVal = SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&)
        If Not RetVal = LB_ERR Then
            RetVal = SendMessage(ListBoxHandle, LB_GETSEL, RetVal, ByVal 0&)
            If RetVal > 0 Then SelCount = 1
        End If
    End If
End If
End Property

Public Property Get Selected(ByVal Index As Long) As Boolean
Attribute Selected.VB_Description = "Returns/sets the selection status of an item."
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Index, ByVal 0&) = LB_ERR Then
        Selected = CBool(SendMessage(ListBoxHandle, LB_GETSEL, Index, ByVal 0&) > 0)
    Else
        Err.Raise 381
    End If
End If
End Property

Public Property Let Selected(ByVal Index As Long, ByVal Value As Boolean)
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Index, ByVal 0&) = LB_ERR Then
        Dim Changed As Boolean, RetVal As Long
        If PropMultiSelect <> vbMultiSelectNone Then
            RetVal = IIf(SendMessage(ListBoxHandle, LB_GETSEL, Index, ByVal 0&) > 0, 1, 0)
            SendMessage ListBoxHandle, LB_SETSEL, IIf(Value = True, 1, 0), ByVal Index
            Changed = CBool(IIf(SendMessage(ListBoxHandle, LB_GETSEL, Index, ByVal 0&) > 0, 1, 0) <> RetVal)
        Else
            RetVal = SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&)
            If Value = False Then
                If SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&) = Index Then
                    If SendMessage(ListBoxHandle, LB_GETSEL, Index, ByVal 0&) > 0 Then SendMessage ListBoxHandle, LB_SETCURSEL, -1, ByVal 0&
                End If
            Else
                SendMessage ListBoxHandle, LB_SETCURSEL, Index, ByVal 0&
            End If
            Changed = CBool(SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&) <> RetVal)
        End If
        If Changed = True Then RaiseEvent Click
    Else
        Err.Raise 381
    End If
End If
End Property

Public Sub SetSelRange(ByVal StartIndex As Long, ByVal EndIndex As Long)
Attribute SetSelRange.VB_Description = "Sets the start and end item for the current selection range."
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, StartIndex, ByVal 0&) = LB_ERR And Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, EndIndex, ByVal 0&) = LB_ERR Then
        Dim RetVal As Long
        RetVal = SendMessage(ListBoxHandle, LB_GETSELCOUNT, 0, ByVal 0&)
        If Not RetVal = LB_ERR Then
            Dim Changed As Boolean
            SendMessage ListBoxHandle, LB_SELITEMRANGEEX, StartIndex, ByVal EndIndex
            Changed = CBool(SendMessage(ListBoxHandle, LB_GETSELCOUNT, 0, ByVal 0&) <> RetVal)
            If Changed = True Then RaiseEvent Click
        Else
            Me.ListIndex = StartIndex
        End If
    Else
        Err.Raise 381
    End If
End If
End Sub

Public Property Get ItemHeight(Optional ByVal Index As Long) As Single
Attribute ItemHeight.VB_Description = "Returns/sets the height of an item. The optional index argument can be specified in an variable owner-drawn list box."
Attribute ItemHeight.VB_MemberFlags = "400"
If ListBoxHandle <> 0 Then
    Dim RetVal As Long
    If PropDrawMode <> LstDrawModeOwnerDrawVariable Then
        If Index = 0 Then
            RetVal = SendMessage(ListBoxHandle, LB_GETITEMHEIGHT, 0, ByVal 0&)
        Else
            RetVal = LB_ERR
        End If
    Else
        RetVal = SendMessage(ListBoxHandle, LB_GETITEMHEIGHT, Index, ByVal 0&)
    End If
    If Not RetVal = LB_ERR Then
        ItemHeight = UserControl.ScaleY(RetVal, vbPixels, vbContainerSize)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Let ItemHeight(Optional ByVal Index As Long, ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If ListBoxHandle <> 0 Then
    Dim RetVal As Long
    If PropDrawMode <> LstDrawModeOwnerDrawVariable Then
        If Index = 0 Then
            RetVal = SendMessage(ListBoxHandle, LB_SETITEMHEIGHT, 0, ByVal CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels)))
        Else
            RetVal = LB_ERR
        End If
    Else
        RetVal = SendMessage(ListBoxHandle, LB_SETITEMHEIGHT, Index, ByVal CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels)))
    End If
    If Not RetVal = LB_ERR Then
        If PropIntegralHeight = True Then
            With UserControl
            MoveWindow ListBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight + 10, 0
            MoveWindow ListBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 0
            End With
            Call UserControl_Resize
        End If
        Me.Refresh
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Get NewIndex() As Long
Attribute NewIndex.VB_Description = "Returns the index of the item most recently added to a control."
Attribute NewIndex.VB_MemberFlags = "400"
NewIndex = ListBoxNewIndex
End Property

Public Property Get TopIndex() As Long
Attribute TopIndex.VB_Description = "Returns/sets which item in a control is displayed in the topmost position."
Attribute TopIndex.VB_MemberFlags = "400"
If ListBoxHandle <> 0 Then TopIndex = SendMessage(ListBoxHandle, LB_GETTOPINDEX, 0, ByVal 0&)
End Property

Public Property Let TopIndex(ByVal Value As Long)
If ListBoxHandle <> 0 Then
    If Value >= 0 Then
        If SendMessage(ListBoxHandle, LB_SETTOPINDEX, Value, ByVal 0&) = LB_ERR Then Err.Raise 380
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get AnchorIndex() As Long
Attribute AnchorIndex.VB_Description = "Returns/sets the anchor item. That is the item from which a multiple selection starts."
Attribute AnchorIndex.VB_MemberFlags = "400"
If ListBoxHandle <> 0 Then AnchorIndex = SendMessage(ListBoxHandle, LB_GETANCHORINDEX, 0, ByVal 0&)
End Property

Public Property Let AnchorIndex(ByVal Value As Long)
If ListBoxHandle <> 0 Then
    If Value < -1 Then
        Err.Raise 380
    Else
        If SendMessage(ListBoxHandle, LB_SETANCHORINDEX, Value, ByVal 0&) = LB_ERR Then Err.Raise 380
    End If
End If
End Property

Public Sub SetColumnWidth(ByVal Value As Single)
Attribute SetColumnWidth.VB_Description = "Sets the width of all columns in a multiple-column list box."
If Value < 0 Then Err.Raise 380
If ListBoxHandle <> 0 Then
    Dim LngValue As Long
    LngValue = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    If LngValue > 0 Then
        SendMessage ListBoxHandle, LB_SETCOLUMNWIDTH, LngValue, ByVal 0&
    Else
        Err.Raise 380
    End If
End If
End Sub

Public Function ItemsPerColumn() As Long
Attribute ItemsPerColumn.VB_Description = "Retrieves the number of items per column."
If ListBoxHandle <> 0 Then ItemsPerColumn = SendMessage(ListBoxHandle, LB_GETLISTBOXINFO, 0, ByVal 0&)
End Function

Public Function SelectedIndices() As Collection
Attribute SelectedIndices.VB_Description = "Returns a reference to a collection containing the indexes to the selected items."
If ListBoxHandle <> 0 Then
    Set SelectedIndices = New Collection
    Dim Count As Long
    Count = SendMessage(ListBoxHandle, LB_GETCOUNT, 0, ByVal 0&)
    If Count > 0 Then
        Dim LngArr() As Long, RetVal As Long
        ReDim LngArr(1 To Count) As Long
        RetVal = SendMessage(ListBoxHandle, LB_GETSELITEMS, Count, ByVal VarPtr(LngArr(1)))
        If Not RetVal = LB_ERR Then
            Dim i As Long
            For i = 1 To RetVal
                SelectedIndices.Add LngArr(i)
            Next i
        Else
            RetVal = SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&)
            If Not RetVal = LB_ERR Then
                If SendMessage(ListBoxHandle, LB_GETSEL, RetVal, ByVal 0&) > 0 Then
                    SelectedIndices.Add RetVal
                End If
            End If
        End If
    End If
End If
End Function

Public Function CheckedIndices() As Collection
Attribute CheckedIndices.VB_Description = "Returns a reference to a collection containing the indexes to the checked items."
If ListBoxHandle <> 0 Then
    Set CheckedIndices = New Collection
    Dim Count As Long
    Count = SendMessage(ListBoxHandle, LB_GETCOUNT, 0, ByVal 0&)
    If Count > 0 Then
        If PropStyle = LstStyleCheckbox Then
            Dim i As Long
            For i = 1 To UBound(ListBoxItemChecked())
                If ListBoxItemChecked(i) = vbChecked Then CheckedIndices.Add (i - 1)
            Next i
        ElseIf PropStyle = LstStyleOption Then
            If ListBoxOptionIndex > -1 Then CheckedIndices.Add ListBoxOptionIndex
        End If
    End If
End If
End Function

Public Function HitTest(ByVal X As Single, ByVal Y As Single) As Long
Attribute HitTest.VB_Description = "Returns the index of the item located at the coordinates of X and Y."
If ListBoxHandle <> 0 Then
    Dim P As POINTAPI
    P.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    P.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    ClientToScreen ListBoxHandle, P
    HitTest = LBItemFromPt(ListBoxHandle, P.X, P.Y, 0)
End If
End Function

Public Function HitTestInsertMark(ByVal X As Single, ByVal Y As Single, Optional ByRef After As Boolean) As Long
Attribute HitTestInsertMark.VB_Description = "Returns the index of the item located at the coordinates of X and Y and retrieves a value that determines where the insertion point should appear."
If ListBoxHandle <> 0 Then
    Dim P As POINTAPI, Index As Long
    P.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    P.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    ClientToScreen ListBoxHandle, P
    Index = LBItemFromPt(ListBoxHandle, P.X, P.Y, 0)
    If Index > -1 Then
        Dim RC As RECT
        SendMessage ListBoxHandle, LB_GETITEMRECT, Index, ByVal VarPtr(RC)
        After = CBool(CLng(UserControl.ScaleY(Y, vbContainerPosition, vbPixels)) > (RC.Top + ((RC.Bottom - RC.Top) / 2)))
    End If
    HitTestInsertMark = Index
End If
End Function

Public Function FindItem(ByVal Text As String, Optional ByVal Index As Long = -1, Optional ByVal Partial As Boolean) As Long
Attribute FindItem.VB_Description = "Finds an item in the list box and returns the index of that item."
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Index, ByVal 0&) = LB_ERR Or Index = -1 Then
        If Partial = True Then
            FindItem = SendMessage(ListBoxHandle, LB_FINDSTRING, Index, ByVal StrPtr(Text))
        Else
            FindItem = SendMessage(ListBoxHandle, LB_FINDSTRINGEXACT, Index, ByVal StrPtr(Text))
        End If
    Else
        Err.Raise 381
    End If
End If
End Function

Public Property Get InsertMark(Optional ByRef After As Boolean) As Long
Attribute InsertMark.VB_Description = "Returns/sets the index of the item where an insertion mark is positioned."
Attribute InsertMark.VB_MemberFlags = "400"
InsertMark = ListBoxInsertMark
After = ListBoxInsertMarkAfter
End Property

Public Property Let InsertMark(Optional ByRef After As Boolean, ByVal Value As Long)
If ListBoxInsertMark = Value And ListBoxInsertMarkAfter = After Then Exit Property
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Value, ByVal 0&) = LB_ERR Or Value = -1 Then
        If ListBoxInsertMark > -1 Then Call InvalidateInsertMark
        ListBoxInsertMark = Value
        ListBoxInsertMarkAfter = After
        If ListBoxInsertMark > -1 Then Call InvalidateInsertMark
    Else
        Err.Raise 381
    End If
End If
End Property

Public Property Get OptionIndex() As Long
Attribute OptionIndex.VB_Description = "Returns/sets the index of the checked item when the style property is set to 2 (option)."
Attribute OptionIndex.VB_MemberFlags = "400"
OptionIndex = ListBoxOptionIndex
End Property

Public Property Let OptionIndex(ByVal Value As Long)
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Value, ByVal 0&) = LB_ERR Or Value = -1 Then
        If PropStyle = LstStyleOption Then
            If Value > -1 Then
                Me.ItemChecked(Value) = True
            Else
                If ListBoxOptionIndex > -1 Then Me.ItemChecked(ListBoxOptionIndex) = False
            End If
        End If
    Else
        Err.Raise 381
    End If
End If
End Property

Public Property Get OLEDraggedItem() As Long
Attribute OLEDraggedItem.VB_Description = "Returns the index of the currently dragged item during an OLE drag/drop operation."
Attribute OLEDraggedItem.VB_MemberFlags = "400"
OLEDraggedItem = ListBoxDragIndex - 1
End Property

Public Function GetIdealHorizontalExtent() As Single
Attribute GetIdealHorizontalExtent.VB_Description = "Gets the ideal value for the horizontal extent property."
If ListBoxHandle <> 0 Then
    Dim Count As Long
    Count = SendMessage(ListBoxHandle, LB_GETCOUNT, 0, ByVal 0&)
    If Count > 0 Then
        Dim RC(0 To 1) As RECT, CX As Long, ScrollWidth As Long, hDC As Long, i As Long, Length As Long, Text As String, Size As SIZEAPI
        GetWindowRect ListBoxHandle, RC(0)
        GetClientRect ListBoxHandle, RC(1)
        If (GetWindowLong(ListBoxHandle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL Then
            Const SM_CXVSCROLL As Long = 2
            ScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
        End If
        hDC = GetDC(ListBoxHandle)
        SelectObject hDC, ListBoxFontHandle
        For i = 0 To Count - 1
            Length = SendMessage(ListBoxHandle, LB_GETTEXTLEN, i, ByVal 0&)
            If Not Length = LB_ERR Then
                Text = String(Length, vbNullChar)
                SendMessage ListBoxHandle, LB_GETTEXT, i, ByVal StrPtr(Text)
                GetTextExtentPoint32 hDC, ByVal StrPtr(Text), Length, Size
                If (Size.CX - ScrollWidth) > CX Then CX = (Size.CX - ScrollWidth)
            End If
        Next i
        ReleaseDC ListBoxHandle, hDC
        If CX > 0 Then GetIdealHorizontalExtent = UserControl.ScaleX(CX + ((RC(0).Right - RC(0).Left) - (RC(1).Right - RC(1).Left)), vbPixels, vbContainerSize)
    End If
End If
End Function

Public Function SelectItem(ByVal Text As String, Optional ByVal Index As Long = -1) As Long
Attribute SelectItem.VB_Description = "Searches for an item that begins with the characters in a specified string. If a matching item is found, the item is selected. The search is not case sensitive."
If ListBoxHandle <> 0 Then
    If Not SendMessage(ListBoxHandle, LB_GETTEXTLEN, Index, ByVal 0&) = LB_ERR Or Index = -1 Then
        SelectItem = SendMessage(ListBoxHandle, LB_SELECTSTRING, Index, ByVal StrPtr(Text))
    Else
        Err.Raise 381
    End If
End If
End Function

Private Sub SetItemCheck(Optional ByVal Index As Long = LB_ERR)
If ListBoxHandle <> 0 Then
    If Index = LB_ERR Then Index = SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&)
    If Not Index = LB_ERR Then
        If Index <= (ListBoxItemCheckedCount - 1) Then
            Dim Changed As Boolean
            If PropStyle = LstStyleCheckbox Then
                Changed = True
            ElseIf PropStyle = LstStyleOption Then
                Changed = CBool(ListBoxOptionIndex <> Index)
            End If
            If Changed = True Then
                Dim Cancel As Boolean
                RaiseEvent ItemBeforeCheck(Index, Cancel)
                If Cancel = False Then
                    Dim RC As RECT
                    If PropStyle = LstStyleCheckbox Then
                        Select Case ListBoxItemChecked(Index + 1)
                            Case vbChecked
                                ListBoxItemChecked(Index + 1) = vbUnchecked
                            Case Else
                                ListBoxItemChecked(Index + 1) = vbChecked
                        End Select
                    ElseIf PropStyle = LstStyleOption Then
                        If ListBoxOptionIndex > -1 Then
                            SendMessage ListBoxHandle, LB_GETITEMRECT, ListBoxOptionIndex, ByVal VarPtr(RC)
                            InvalidateRect ListBoxHandle, RC, 0
                        End If
                        ListBoxOptionIndex = Index
                    End If
                    SendMessage ListBoxHandle, LB_GETITEMRECT, Index, ByVal VarPtr(RC)
                    InvalidateRect ListBoxHandle, RC, 0
                    RaiseEvent ItemCheck(Index)
                End If
            End If
        End If
    End If
End If
End Sub

Private Function CheckTopIndex() As Boolean
Dim TopIndex As Long
If ListBoxHandle <> 0 Then TopIndex = SendMessage(ListBoxHandle, LB_GETTOPINDEX, 0, ByVal 0&)
If TopIndex <> ListBoxTopIndex Then
    ListBoxTopIndex = TopIndex
    If ListBoxInsertMark > -1 Then Call InvalidateInsertMark
    RaiseEvent Scroll
    CheckTopIndex = True
End If
End Function

Private Sub InvalidateInsertMark()
If ListBoxHandle <> 0 Then
    If SendMessage(ListBoxHandle, LB_GETTEXTLEN, ListBoxInsertMark, ByVal 0&) = LB_ERR Then Exit Sub
    Dim RC As RECT
    SendMessage ListBoxHandle, LB_GETITEMRECT, ListBoxInsertMark, ByVal VarPtr(RC)
    If ListBoxInsertMarkAfter = False Then
        RC.Bottom = RC.Top + 1
        RC.Top = RC.Top - 1
    Else
        RC.Top = RC.Bottom - 1
        RC.Bottom = RC.Bottom + 1
    End If
    RC.Top = RC.Top - 2
    RC.Bottom = RC.Bottom + 2
    InvalidateRect ListBoxHandle, RC, 1
End If
End Sub

Private Sub DrawInsertMark()
If ListBoxHandle <> 0 Then
    If SendMessage(ListBoxHandle, LB_GETTEXTLEN, ListBoxInsertMark, ByVal 0&) = LB_ERR Then Exit Sub
    Dim RC As RECT, hRgn As Long, hDC As Long, Brush As Long, OldBrush As Long
    GetClientRect ListBoxHandle, RC
    hDC = GetDC(ListBoxHandle)
    If hDC <> 0 Then
        hRgn = CreateRectRgnIndirect(RC)
        If hRgn <> 0 Then ExtSelectClipRgn hDC, hRgn, RGN_COPY
        SendMessage ListBoxHandle, LB_GETITEMRECT, ListBoxInsertMark, ByVal VarPtr(RC)
        If ListBoxInsertMarkAfter = False Then
            RC.Bottom = RC.Top + 1
            RC.Top = RC.Top - 1
        Else
            RC.Top = RC.Bottom - 1
            RC.Bottom = RC.Bottom + 1
        End If
        Brush = CreateSolidBrush(WinColor(PropInsertMarkColor))
        If Brush <> 0 Then OldBrush = SelectObject(hDC, Brush)
        PatBlt hDC, RC.Left, RC.Top - 2, 1, 6, vbPatCopy
        PatBlt hDC, RC.Left + 1, RC.Top - 1, 1, 4, vbPatCopy
        PatBlt hDC, RC.Left + 2, RC.Top, RC.Right - RC.Left - 2, RC.Bottom - RC.Top, vbPatCopy
        PatBlt hDC, RC.Right - 2, RC.Top - 1, 1, 4, vbPatCopy
        PatBlt hDC, RC.Right - 1, RC.Top - 2, 1, 6, vbPatCopy
        If OldBrush <> 0 Then SelectObject hDC, OldBrush
        If Brush <> 0 Then DeleteObject Brush
        If hRgn <> 0 Then
            ExtSelectClipRgn hDC, 0, RGN_COPY
            DeleteObject hRgn
        End If
        ReleaseDC ListBoxHandle, hDC
    End If
End If
End Sub

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcUserControlDesignMode(hWnd, wMsg, wParam, lParam)
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
                If PropStyle <> LstStyleStandard And KeyCode = vbKeySpace Then Call SetItemCheck
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            ListBoxCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If ListBoxCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(ListBoxCharCodeCache And &HFFFF&)
            ListBoxCharCodeCache = 0
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
        If PropOLEDragMode = vbOLEDragAutomatic Or PropStyle <> LstStyleStandard Then
            Dim P1 As POINTAPI, P2 As POINTAPI, Index As Long
            P1.X = Get_X_lParam(lParam)
            P1.Y = Get_Y_lParam(lParam)
            P2.X = P1.X
            P2.Y = P1.Y
            ClientToScreen ListBoxHandle, P2
            Index = LBItemFromPt(ListBoxHandle, P2.X, P2.Y, 0)
            If Index > -1 Then
                Dim IsItemCheck As Boolean
                If PropStyle <> LstStyleStandard Then
                    If Index <> SendMessage(ListBoxHandle, LB_GETCURSEL, 0, ByVal 0&) Then
                        Dim RC As RECT
                        SendMessage ListBoxHandle, LB_GETITEMRECT, Index, ByVal VarPtr(RC)
                        If PropRightToLeft = False Then
                            IsItemCheck = CBool(Get_X_lParam(lParam) < (RC.Left + ListBoxStateImageSize))
                        Else
                            IsItemCheck = CBool(Get_X_lParam(lParam) >= (RC.Right - ListBoxStateImageSize))
                        End If
                    Else
                        IsItemCheck = True
                    End If
                End If
                If PropOLEDragMode = vbOLEDragAutomatic Then
                    If SendMessage(ListBoxHandle, LB_GETSEL, Index, ByVal 0&) > 0 Then
                        If GetFocus() <> hWnd Then SetFocusAPI UserControl.hWnd ' UCNoSetFocusFwd not applicable
                        RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(P1.X, vbPixels, vbTwips), UserControl.ScaleY(P1.Y, vbPixels, vbTwips))
                        If DragDetect(ListBoxHandle, CUIntToInt(P2.X And &HFFFF&), CUIntToInt(P2.Y And &HFFFF&)) <> 0 Then
                            ListBoxDragIndexBuffer = Index + 1
                            Me.OLEDrag
                            ListBoxDragIndexBuffer = 0
                            WindowProcControl = 0
                        Else
                            WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
                            ReleaseCapture
                            If IsItemCheck = True Then Call SetItemCheck(Index)
                            RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(P1.X, vbPixels, vbTwips), UserControl.ScaleY(P1.Y, vbPixels, vbTwips))
                        End If
                        Exit Function
                    End If
                End If
                If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
                If IsItemCheck = True Then
                    WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
                    RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), UserControl.ScaleX(P1.X, vbPixels, vbTwips), UserControl.ScaleY(P1.Y, vbPixels, vbTwips))
                    If IsItemCheck = True Then Call SetItemCheck(Index)
                    Exit Function
                End If
            Else
                If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
            End If
        Else
            If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
        End If
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
    Case WM_CONTEXTMENU
        If wParam = ListBoxHandle Then
            Dim P3 As POINTAPI
            P3.X = Get_X_lParam(lParam)
            P3.Y = Get_Y_lParam(lParam)
            If P3.X = -1 And P3.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(-1, -1)
            Else
                ScreenToClient ListBoxHandle, P3
                RaiseEvent ContextMenu(UserControl.ScaleX(P3.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P3.Y, vbPixels, vbContainerPosition))
            End If
        End If
    Case WM_HSCROLL, WM_VSCROLL
        If Not (wMsg = WM_HSCROLL And PropMultiColumn = False) Then
            Select Case LoWord(wParam)
                Case SB_THUMBPOSITION, SB_THUMBTRACK
                    ' HiWord carries only 16 bits of scroll box position data.
                    ' Below workaround will circumvent the 16-bit barrier by using the 32-bit GetScrollInfo function.
                    Dim dwStyle As Long
                    dwStyle = GetWindowLong(ListBoxHandle, GWL_STYLE)
                    If lParam = 0 And ((wMsg = WM_HSCROLL And (dwStyle And WS_HSCROLL) = WS_HSCROLL) Or (wMsg = WM_VSCROLL And (dwStyle And WS_VSCROLL) = WS_VSCROLL)) Then
                        Dim SCI As SCROLLINFO, wBar As Long, PrevPos As Long
                        SCI.cbSize = LenB(SCI)
                        SCI.fMask = SIF_RANGE Or SIF_POS Or SIF_TRACKPOS
                        If wMsg = WM_HSCROLL Then
                            wBar = SB_HORZ
                        ElseIf wMsg = WM_VSCROLL Then
                            wBar = SB_VERT
                        End If
                        GetScrollInfo ListBoxHandle, wBar, SCI
                        PrevPos = SCI.nPos
                        Select Case LoWord(wParam)
                            Case SB_THUMBPOSITION
                                SCI.nPos = SCI.nTrackPos
                            Case SB_THUMBTRACK
                                If PropScrollTrack = True Then SCI.nPos = SCI.nTrackPos
                        End Select
                        If PrevPos <> SCI.nPos Then
                            If wMsg = WM_HSCROLL And PropMultiColumn = True Then
                                If (GetWindowLong(ListBoxHandle, GWL_EXSTYLE) And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then SCI.nPos = (((SCI.nMax - SCI.nMin) - 1) - SCI.nPos)
                                SCI.nPos = SCI.nPos * Me.ItemsPerColumn
                            End If
                            ' SetScrollInfo function not needed as LB_SETTOPINDEX itself will do the scrolling.
                            SendMessage ListBoxHandle, LB_SETTOPINDEX, SCI.nPos, ByVal 0&
                        End If
                        WindowProcControl = 0
                        Exit Function
                    End If
            End Select
        End If
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
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MOUSEMOVE
                If (GetMouseStateFromParam(wParam) And vbLeftButton) = vbLeftButton Then
                    If CheckTopIndex() = False And ListBoxInsertMark > -1 Then Call InvalidateInsertMark
                End If
                If ListBoxMouseOver = False And PropMouseTrack = True Then
                    ListBoxMouseOver = True
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
        End Select
    Case WM_MOUSELEAVE
        If ListBoxMouseOver = True Then
            ListBoxMouseOver = False
            RaiseEvent MouseLeave
        End If
    Case WM_MOUSEWHEEL, WM_HSCROLL, WM_VSCROLL, LB_SETTOPINDEX
        If CheckTopIndex() = False And ListBoxInsertMark > -1 Then Call InvalidateInsertMark
    Case WM_PAINT
        If ListBoxInsertMark > -1 Then Call DrawInsertMark
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        If lParam = ListBoxHandle Then
            Select Case HiWord(wParam)
                Case LBN_SELCHANGE
                    If CheckTopIndex() = False And ListBoxInsertMark > -1 Then Call InvalidateInsertMark
                    RaiseEvent Click
                Case LBN_SELCANCEL
                    If ListBoxInsertMark > -1 Then Call InvalidateInsertMark
                    RaiseEvent Click
                Case LBN_DBLCLK
                    RaiseEvent DblClick
            End Select
        End If
    Case WM_MEASUREITEM
        If PropDrawMode = LstDrawModeOwnerDrawVariable Then
            Dim MIS As MEASUREITEMSTRUCT
            CopyMemory MIS, ByVal lParam, LenB(MIS)
            If MIS.CtlType = ODT_LISTBOX And MIS.ItemID > -1 Then
                With MIS
                RaiseEvent ItemMeasure(.ItemID, .ItemHeight)
                End With
                CopyMemory ByVal lParam, MIS, LenB(MIS)
                WindowProcUserControl = 1
                Exit Function
            End If
        End If
    Case WM_DRAWITEM
        Dim DIS As DRAWITEMSTRUCT
        CopyMemory DIS, ByVal lParam, LenB(DIS)
        If DIS.CtlType = ODT_LISTBOX And DIS.hWndItem = ListBoxHandle And DIS.ItemID > -1 Then
            If PropStyle <> LstStyleStandard Then
                Dim BackColorBrush As Long, BackColorSelBrush As Long
                BackColorBrush = CreateSolidBrush(WinColor(Me.BackColor))
                If (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED And PropAllowSelection = True Then BackColorSelBrush = CreateSolidBrush(WinColor(vbHighlight))
                Dim RC As RECT
                With DIS.RCItem
                If PropRightToLeft = False Then
                    SetRect RC, .Left + 1, .Top + 1, .Left + ListBoxStateImageSize - 1, .Bottom - 1
                    .Left = .Left + ListBoxStateImageSize
                Else
                    SetRect RC, .Right - ListBoxStateImageSize + 1, .Top + 1, .Right - 1, .Bottom - 1
                    .Right = .Right - ListBoxStateImageSize
                End If
                End With
                If BackColorSelBrush <> 0 Then
                    FillRect DIS.hDC, DIS.RCItem, BackColorSelBrush
                    DeleteObject BackColorSelBrush
                Else
                    FillRect DIS.hDC, DIS.RCItem, BackColorBrush
                End If
                FillRect DIS.hDC, RC, BackColorBrush
                DeleteObject BackColorBrush
                
                #If ImplementThemedButton = True Then
                
                Dim Theme As Long
                If EnabledVisualStyles() = True And PropVisualStyles = True Then Theme = OpenThemeData(ListBoxHandle, StrPtr("Button"))
                If Theme <> 0 Then
                    Dim ButtonPart As Long, CheckState As Long
                    If PropStyle = LstStyleCheckbox Then
                        ButtonPart = BP_CHECKBOX
                        If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                            CheckState = CBS_UNCHECKEDNORMAL
                        Else
                            CheckState = CBS_UNCHECKEDDISABLED
                        End If
                        If DIS.ItemID <= (ListBoxItemCheckedCount - 1) Then
                            If ListBoxItemChecked(DIS.ItemID + 1) = vbChecked Then
                                If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                                    CheckState = CBS_CHECKEDNORMAL
                                Else
                                    CheckState = CBS_CHECKEDDISABLED
                                End If
                            End If
                        End If
                    ElseIf PropStyle = LstStyleOption Then
                        ButtonPart = BP_RADIOBUTTON
                        If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                            CheckState = RBS_UNCHECKEDNORMAL
                        Else
                            CheckState = RBS_UNCHECKEDDISABLED
                        End If
                        If DIS.ItemID <= (ListBoxItemCheckedCount - 1) Then
                            If ListBoxOptionIndex = DIS.ItemID Then
                                If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                                    CheckState = CBS_CHECKEDNORMAL
                                Else
                                    CheckState = CBS_CHECKEDDISABLED
                                End If
                            End If
                        End If
                    End If
                    If IsThemeBackgroundPartiallyTransparent(Theme, ButtonPart, CheckState) <> 0 Then DrawThemeParentBackground DIS.hWndItem, DIS.hDC, RC
                    DrawThemeBackground Theme, DIS.hDC, ButtonPart, CheckState, RC, RC
                    CloseThemeData Theme
                Else
                    Dim Flags As Long
                    Flags = DFCS_FLAT
                    If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then Flags = Flags Or DFCS_INACTIVE
                    If PropStyle = LstStyleCheckbox Then
                        Flags = Flags Or DFCS_BUTTONCHECK
                        If DIS.ItemID <= (ListBoxItemCheckedCount - 1) Then
                            If ListBoxItemChecked(DIS.ItemID + 1) = vbChecked Then Flags = Flags Or DFCS_CHECKED
                        End If
                    ElseIf PropStyle = LstStyleOption Then
                        Flags = Flags Or DFCS_BUTTONRADIO
                        If DIS.ItemID <= (ListBoxItemCheckedCount - 1) Then
                            If ListBoxOptionIndex = DIS.ItemID Then Flags = Flags Or DFCS_CHECKED
                        End If
                    End If
                    DrawFrameControl DIS.hDC, RC, DFC_BUTTON, Flags
                End If
                
                #Else
                
                Dim Flags As Long
                Flags = DFCS_FLAT
                If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then Flags = Flags Or DFCS_INACTIVE
                If PropStyle = LstStyleCheckbox Then
                    Flags = Flags Or DFCS_BUTTONCHECK
                    If DIS.ItemID <= (ListBoxItemCheckedCount - 1) Then
                        If ListBoxItemChecked(DIS.ItemID + 1) = vbChecked Then Flags = Flags Or DFCS_CHECKED
                    End If
                ElseIf PropStyle = LstStyleOption Then
                    Flags = Flags Or DFCS_BUTTONRADIO
                    If DIS.ItemID <= (ListBoxItemCheckedCount - 1) Then
                        If ListBoxOptionIndex = DIS.ItemID Then Flags = Flags Or DFCS_CHECKED
                    End If
                End If
                DrawFrameControl DIS.hDC, RC, DFC_BUTTON, Flags
                
                #End If
                
                Dim Length As Long
                Length = SendMessage(ListBoxHandle, LB_GETTEXTLEN, DIS.ItemID, ByVal 0&)
                If Not Length = LB_ERR Then
                    Dim Text As String
                    Text = String(Length, vbNullChar)
                    SendMessage ListBoxHandle, LB_GETTEXT, DIS.ItemID, ByVal StrPtr(Text)
                    Dim OldTextAlign As Long, OldBkMode As Long, OldTextColor As Long
                    If PropRightToLeft = True Then OldTextAlign = SetTextAlign(DIS.hDC, TA_RTLREADING Or TA_RIGHT)
                    OldBkMode = SetBkMode(DIS.hDC, 1)
                    If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                        OldTextColor = SetTextColor(DIS.hDC, WinColor(vbGrayText))
                    ElseIf (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED And PropAllowSelection = True Then
                        OldTextColor = SetTextColor(DIS.hDC, WinColor(vbHighlightText))
                    Else
                        OldTextColor = SetTextColor(DIS.hDC, WinColor(Me.ForeColor))
                    End If
                    If PropRightToLeft = False Then
                        If PropUseTabStops = False Then
                            TextOut DIS.hDC, DIS.RCItem.Left + (1 * PixelsPerDIP_X()), DIS.RCItem.Top, StrPtr(Text), Length
                        Else
                            TabbedTextOut DIS.hDC, DIS.RCItem.Left + (1 * PixelsPerDIP_X()), DIS.RCItem.Top, StrPtr(Text), Len(Text), 0, 0, 0
                        End If
                    Else
                        If PropUseTabStops = False Then
                            TextOut DIS.hDC, DIS.RCItem.Right - (1 * PixelsPerDIP_X()), DIS.RCItem.Top, StrPtr(Text), Length
                        Else
                            TabbedTextOut DIS.hDC, DIS.RCItem.Right - (1 * PixelsPerDIP_X()), DIS.RCItem.Top, StrPtr(Text), Len(Text), 0, 0, 0
                        End If
                    End If
                    SetBkMode DIS.hDC, OldBkMode
                    SetTextColor DIS.hDC, OldTextColor
                    If PropRightToLeft = True Then SetTextAlign DIS.hDC, OldTextAlign
                End If
                If (DIS.ItemState And ODS_FOCUS) = ODS_FOCUS Then DrawFocusRect DIS.hDC, DIS.RCItem
            Else
                With DIS
                RaiseEvent ItemDraw(.ItemID, .ItemAction, .ItemState, .hDC, .RCItem.Left, .RCItem.Top, .RCItem.Right, .RCItem.Bottom)
                End With
            End If
            WindowProcUserControl = 1
            Exit Function
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI ListBoxHandle
End Function

Private Function WindowProcUserControlDesignMode(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg = WM_DRAWITEM Then
    WindowProcUserControlDesignMode = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Exit Function
End If
WindowProcUserControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
