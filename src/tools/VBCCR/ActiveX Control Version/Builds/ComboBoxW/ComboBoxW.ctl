VERSION 5.00
Begin VB.UserControl ComboBoxW 
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   PropertyPages   =   "ComboBoxW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ComboBoxW.ctx":0035
End
Attribute VB_Name = "ComboBoxW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
#If False Then
Private CboStyleDropDownCombo, CboStyleSimpleCombo, CboStyleDropDownList
Private CboCharacterCasingNormal, CboCharacterCasingUpper, CboCharacterCasingLower
Private CboDrawModeNormal, CboDrawModeOwnerDrawFixed, CboDrawModeOwnerDrawVariable
#End If
Public Enum CboStyleConstants
CboStyleDropDownCombo = 0
CboStyleSimpleCombo = 1
CboStyleDropDownList = 2
End Enum
Public Enum CboCharacterCasingConstants
CboCharacterCasingNormal = 0
CboCharacterCasingUpper = 1
CboCharacterCasingLower = 2
End Enum
Public Enum CboDrawModeConstants
CboDrawModeNormal = 0
CboDrawModeOwnerDrawFixed = 1
CboDrawModeOwnerDrawVariable = 2
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
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Private Type MEASUREITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
ItemWidth As Long
ItemHeight As Long
ItemData As LongPtr
End Type
Private Type DRAWITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
ItemAction As Long
ItemState As Long
hWndItem As LongPtr
hDC As LongPtr
RCItem As RECT
ItemData As LongPtr
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
Private Type COMBOBOXINFO
cbSize As Long
RCItem As RECT
RCButton As RECT
StateButton As Long
hWndCombo As LongPtr
hWndItem As LongPtr
hWndList As LongPtr
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Scroll()
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Public Event ContextMenu(ByRef Handled As Boolean, ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event DropDown()
Attribute DropDown.VB_Description = "Occurs when the drop-down list is about to drop down."
Public Event CloseUp()
Attribute CloseUp.VB_Description = "Occurs when the drop-down list has been closed."
Public Event ItemMeasure(ByVal Item As Long, ByRef ItemHeight As Long)
Attribute ItemMeasure.VB_Description = "Occurs each time an variable owner-drawn combo box item's size needs to be determined in preparation of drawing it."
Public Event ItemDraw(ByVal Item As Long, ByVal ItemAction As Long, ByVal ItemState As Long, ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Attribute ItemDraw.VB_Description = "Occurs when a visual aspect of an owner-drawn combo box has changed."
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
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function GetComboBoxInfo Lib "user32" (ByVal hWndCombo As LongPtr, ByRef CBI As COMBOBOXINFO) As Long
Private Declare PtrSafe Function LBItemFromPt Lib "comctl32" (ByVal hLB As LongPtr, ByVal XY As Currency, ByVal bAutoScroll As Long) As Long
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function MapWindowPoints Lib "user32" (ByVal hWndFrom As LongPtr, ByVal hWndTo As LongPtr, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As LongPtr, ByVal lpsz As LongPtr, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hDC As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nBkMode As Long) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetScrollInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Private Declare PtrSafe Function DragDetect Lib "user32" (ByVal hWnd As LongPtr, ByVal XY As Currency) As Long
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetMessagePos Lib "user32" () As Long
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal XY As Currency) As Long
Private Declare PtrSafe Function GetCursor Lib "user32" () As LongPtr
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hWndCombo As Long, ByRef CBI As COMBOBOXINFO) As Long
Private Declare Function LBItemFromPt Lib "comctl32" (ByVal hLB As Long, ByVal XY As Currency, ByVal bAutoScroll As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
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
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function DragDetect Lib "user32" (ByVal hWnd As Long, ByVal XY As Currency) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal XY As Currency) As Long
Private Declare Function GetCursor Lib "user32" () As Long
#End If
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
#If VBA7 Then
Private Const HWND_DESKTOP As LongPtr = &H0
#Else
Private Const HWND_DESKTOP As Long = &H0
#End If
Private Const GWL_STYLE As Long = (-16)
Private Const CF_UNICODETEXT As Long = 13
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const SW_HIDE As Long = &H0
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
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
Private Const WM_SIZE As Long = &H5
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_COMMAND As Long = &H111
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_DRAWITEM As Long = &H2B, ODT_COMBOBOX As Long = &H3
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_VSCROLL As Long = &H115
Private Const SB_VERT As Long = 1
Private Const SB_THUMBPOSITION As Long = 4, SB_THUMBTRACK As Long = 5
Private Const SIF_POS As Long = &H4
Private Const SIF_TRACKPOS As Long = &H10
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_CTLCOLORLISTBOX As Long = &H134
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const WM_CLEAR As Long = &H303
Private Const EM_SETREADONLY As Long = &HCF
Private Const EM_SETSEL As Long = &HB1
Private Const EM_REPLACESEL As Long = &HC2
Private Const LB_ERR As Long = (-1)
Private Const LB_SETTOPINDEX As Long = &H197
Private Const CB_ERR As Long = (-1)
Private Const CB_LIMITTEXT As Long = &H141
Private Const CB_ADDSTRING As Long = &H143
Private Const CB_DELETESTRING As Long = &H144
Private Const CB_GETCOUNT As Long = &H146
Private Const CB_GETCURSEL As Long = &H147
Private Const CB_INSERTSTRING As Long = &H14A
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_SETCURSEL As Long = &H14E
Private Const CB_GETDROPPEDCONTROLRECT As Long = &H152
Private Const CB_GETTOPINDEX As Long = &H15B
Private Const CB_SETTOPINDEX As Long = &H15C
Private Const CB_GETHORIZONTALEXTENT As Long = &H15D
Private Const CB_SETHORIZONTALEXTENT As Long = &H15E
Private Const CB_GETDROPPEDWIDTH As Long = &H15F
Private Const CB_SETDROPPEDWIDTH As Long = &H160
Private Const CB_GETLBTEXT As Long = &H148
Private Const CB_GETLBTEXTLEN As Long = &H149
Private Const CB_GETEDITSEL As Long = &H140
Private Const CB_SETEDITSEL As Long = &H142
Private Const CB_RESETCONTENT As Long = &H14B
Private Const CB_SELECTSTRING As Long = &H14D
Private Const CB_SETITEMHEIGHT As Long = &H153
Private Const CB_GETITEMHEIGHT As Long = &H154
Private Const CB_GETDROPPEDSTATE As Long = &H157
Private Const CB_GETCOMBOBOXINFO As Long = &H164 ' Unsupported on W2K
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const CB_GETITEMDATA As Long = &H150
Private Const CB_SETITEMDATA As Long = &H151
Private Const CB_SETEXTENDEDUI As Long = &H155
Private Const CB_GETEXTENDEDUI As Long = &H156
Private Const CB_FINDSTRINGEXACT As Long = &H158
Private Const CBM_FIRST As Long = &H1700
Private Const CB_SETMINVISIBLE As Long = (CBM_FIRST + 1)
Private Const CB_GETMINVISIBLE As Long = (CBM_FIRST + 2)
Private Const CB_SETCUEBANNER As Long = (CBM_FIRST + 3)
Private Const CB_GETCUEBANNER As Long = (CBM_FIRST + 4)
Private Const EM_GETSEL As Long = &HB0
Private Const EM_POSFROMCHAR As Long = &HD6
Private Const EM_CHARFROMPOS As Long = &HD7
Private Const CBS_AUTOHSCROLL As Long = &H40
Private Const CBS_SIMPLE As Long = &H1
Private Const CBS_DROPDOWN As Long = &H2
Private Const CBS_DROPDOWNLIST As Long = &H3
Private Const CBS_OWNERDRAWFIXED As Long = &H10
Private Const CBS_OWNERDRAWVARIABLE As Long = &H20
Private Const CBS_SORT As Long = &H100
Private Const CBS_HASSTRINGS As Long = &H200
Private Const CBS_DISABLENOSCROLL As Long = &H800
Private Const CBS_NOINTEGRALHEIGHT As Long = &H400
Private Const CBS_UPPERCASE As Long = &H2000
Private Const CBS_LOWERCASE As Long = &H4000
Private Const CBN_SELCHANGE As Long = 1
Private Const CBN_DBLCLK As Long = 2
Private Const CBN_EDITCHANGE As Long = 5
Private Const CBN_EDITUPDATE As Long = 6
Private Const CBN_DROPDOWN As Long = 7
Private Const CBN_CLOSEUP As Long = 8
Private Const CBN_SELENDOK As Long = 9
Private Const CBN_SELENDCANCEL As Long = 10
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private ComboBoxHandle As LongPtr, ComboBoxEditHandle As LongPtr, ComboBoxListHandle As LongPtr
Private ComboBoxFontHandle As LongPtr
Private ComboBoxListBackColorBrush As LongPtr
Private ComboBoxIMCHandle As LongPtr
Private ComboBoxCharCodeCache As Long
Private ComboBoxMouseOver(0 To 2) As Boolean
Private ComboBoxDesignMode As Boolean
Private ComboBoxNewIndex As Long
Private ComboBoxTopIndex As Long
Private ComboBoxResizeFrozen As Boolean
Private ComboBoxInitFieldHeight As Long
Private ComboBoxDropDownHeightState As Boolean
Private ComboBoxAutoDragInSel As Boolean, ComboBoxAutoDragIsActive As Boolean
Private ComboBoxAutoDragSelStart As Integer, ComboBoxAutoDragSelEnd As Integer
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropRedraw As Boolean
Private PropStyle As CboStyleConstants
Private PropLocked As Boolean
Private PropText As String
Private PropExtendedUI As Boolean
Private PropMaxDropDownItems As Integer
Private PropIntegralHeight As Boolean
Private PropMaxLength As Long
Private PropCueBanner As String
Private PropUseListBackColor As Boolean
Private PropUseListForeColor As Boolean
Private PropListBackColor As OLE_COLOR
Private PropListForeColor As OLE_COLOR
Private PropSorted As Boolean
Private PropHorizontalExtent As Long
Private PropDisableNoScroll As Boolean
Private PropCharacterCasing As CboCharacterCasingConstants
Private PropDrawMode As CboDrawModeConstants
Private PropIMEMode As CCIMEModeConstants
Private PropScrollTrack As Boolean
Private PropAutoSelect As Boolean

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

#If VBA7 Then
Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal Shift As Long)
#Else
Private Sub IOleInPlaceActiveObjectVB_TranslateAccelerator(ByRef Handled As Boolean, ByRef RetVal As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
#End If
If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
    Dim KeyCode As Integer, IsInputKey As Boolean
    KeyCode = CLng(wParam) And &HFF&
    If wMsg = WM_KEYDOWN Then
        RaiseEvent PreviewKeyDown(KeyCode, IsInputKey)
    ElseIf wMsg = WM_KEYUP Then
        RaiseEvent PreviewKeyUp(KeyCode, IsInputKey)
    End If
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyTab, vbKeyReturn, vbKeyEscape
            If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
                If SendMessage(ComboBoxHandle, CB_GETDROPPEDSTATE, 0, ByVal 0&) = 0 Or PropStyle = CboStyleSimpleCombo Then
                    If IsInputKey = False Then Exit Sub
                Else
                    If PropStyle = CboStyleDropDownCombo Then SendMessage ComboBoxHandle, CB_SHOWDROPDOWN, 0, ByVal 0&
                End If
            ElseIf KeyCode = vbKeyTab Then
                If IsInputKey = False Then Exit Sub
            End If
            SendMessage hWnd, wMsg, wParam, ByVal lParam
            Handled = True
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
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
ComboBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropOLEDragMode = vbOLEDragManual
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropRedraw = True
PropStyle = CboStyleDropDownCombo
PropLocked = False
PropText = Ambient.DisplayName
PropExtendedUI = False
PropMaxDropDownItems = 9
PropIntegralHeight = True
PropMaxLength = 0
PropCueBanner = vbNullString
PropUseListBackColor = False
PropUseListForeColor = False
PropListBackColor = vbWindowBackground
PropListForeColor = vbWindowText
PropSorted = False
PropHorizontalExtent = 0
PropDisableNoScroll = False
PropCharacterCasing = CboCharacterCasingNormal
PropDrawMode = CboDrawModeNormal
PropIMEMode = CCIMEModeNoControl
PropScrollTrack = True
PropAutoSelect = False
Call CreateComboBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
ComboBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
PropOLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
Me.BackColor = .ReadProperty("BackColor", vbWindowBackground)
Me.ForeColor = .ReadProperty("ForeColor", vbWindowText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropRedraw = .ReadProperty("Redraw", True)
PropStyle = .ReadProperty("Style", CboStyleDropDownCombo)
PropLocked = .ReadProperty("Locked", False)
PropText = VarToStr(.ReadProperty("Text", vbNullString))
PropExtendedUI = .ReadProperty("ExtendedUI", False)
PropMaxDropDownItems = .ReadProperty("MaxDropDownItems", 9)
PropIntegralHeight = .ReadProperty("IntegralHeight", True)
PropMaxLength = .ReadProperty("MaxLength", 0)
PropCueBanner = VarToStr(.ReadProperty("CueBanner", vbNullString))
PropUseListBackColor = .ReadProperty("UseListBackColor", False)
PropUseListForeColor = .ReadProperty("UseListForeColor", False)
PropListBackColor = .ReadProperty("ListBackColor", vbWindowBackground)
PropListForeColor = .ReadProperty("ListForeColor", vbWindowText)
PropSorted = .ReadProperty("Sorted", False)
PropHorizontalExtent = .ReadProperty("HorizontalExtent", 0)
PropDisableNoScroll = .ReadProperty("DisableNoScroll", False)
PropCharacterCasing = .ReadProperty("CharacterCasing", CboCharacterCasingNormal)
PropDrawMode = .ReadProperty("DrawMode", CboDrawModeNormal)
PropIMEMode = .ReadProperty("IMEMode", CCIMEModeNoControl)
PropScrollTrack = .ReadProperty("ScrollTrack", True)
PropAutoSelect = .ReadProperty("AutoSelect", False)
End With
Call CreateComboBox
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbWindowBackground
.WriteProperty "ForeColor", Me.ForeColor, vbWindowText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "Redraw", PropRedraw, True
.WriteProperty "Style", PropStyle, CboStyleDropDownCombo
.WriteProperty "Locked", PropLocked, False
.WriteProperty "Text", StrToVar(PropText), vbNullString
.WriteProperty "ExtendedUI", PropExtendedUI, False
.WriteProperty "MaxDropDownItems", PropMaxDropDownItems, 9
.WriteProperty "IntegralHeight", PropIntegralHeight, True
.WriteProperty "MaxLength", PropMaxLength, 0
.WriteProperty "CueBanner", StrToVar(PropCueBanner), vbNullString
.WriteProperty "UseListBackColor", PropUseListBackColor, False
.WriteProperty "UseListForeColor", PropUseListForeColor, False
.WriteProperty "ListBackColor", PropListBackColor, vbWindowBackground
.WriteProperty "ListForeColor", PropListForeColor, vbWindowText
.WriteProperty "Sorted", PropSorted, False
.WriteProperty "HorizontalExtent", PropHorizontalExtent, 0
.WriteProperty "DisableNoScroll", PropDisableNoScroll, False
.WriteProperty "CharacterCasing", PropCharacterCasing, CboCharacterCasingNormal
.WriteProperty "DrawMode", PropDrawMode, CboDrawModeNormal
.WriteProperty "IMEMode", PropIMEMode, CCIMEModeNoControl
.WriteProperty "ScrollTrack", PropScrollTrack, True
.WriteProperty "AutoSelect", PropAutoSelect, False
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
If PropOLEDragMode = vbOLEDragAutomatic And ComboBoxAutoDragIsActive = True And Effect = vbDropEffectMove Then
    If ComboBoxEditHandle <> NULL_PTR Then
        SendMessage ComboBoxEditHandle, EM_SETSEL, ComboBoxAutoDragSelStart, ByVal ComboBoxAutoDragSelEnd
        SendMessage ComboBoxEditHandle, WM_CLEAR, 0, ByVal 0&
    End If
End If
RaiseEvent OLECompleteDrag(Effect)
ComboBoxAutoDragIsActive = False
ComboBoxAutoDragSelStart = 0
ComboBoxAutoDragSelEnd = 0
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
If PropOLEDragMode = vbOLEDragAutomatic Then
    Dim Text As String
    Text = Me.SelText
    Data.SetData StrToVar(Text & vbNullChar), CF_UNICODETEXT
    Data.SetData Text, vbCFText
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    ComboBoxAutoDragIsActive = True
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then ComboBoxAutoDragIsActive = False
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
UserControl.OLEDrag
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If ComboBoxDesignMode = True And PropertyName = "DisplayName" And PropStyle = CboStyleDropDownList Then
    If ComboBoxHandle <> NULL_PTR Then
        If SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&) > 0 Then
            Dim Buffer As String
            Buffer = Ambient.DisplayName
            SendMessage ComboBoxHandle, CB_RESETCONTENT, 0, ByVal 0&
            SendMessage ComboBoxHandle, CB_ADDSTRING, 0, ByVal StrPtr(Buffer)
            SendMessage ComboBoxHandle, CB_SETCURSEL, 0, ByVal 0&
        End If
    End If
End If
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Or ComboBoxResizeFrozen = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If ComboBoxHandle = NULL_PTR Then InProc = False: Exit Sub
Dim WndRect As RECT
If PropStyle <> CboStyleSimpleCombo Then
    If .ScaleHeight > 0 Then MoveWindow ComboBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    GetWindowRect ComboBoxHandle, WndRect
    If (WndRect.Bottom - WndRect.Top) <> .ScaleHeight Or (WndRect.Right - WndRect.Left) <> .ScaleWidth Then
        .Extender.Height = .ScaleY((WndRect.Bottom - WndRect.Top), vbPixels, vbContainerSize)
        If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
    End If
    Call CheckDropDownHeight(True)
Else
    Dim ListRect As RECT, EditHeight As Long, Height As Long
    MoveWindow ComboBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight + IIf(GetIntegralHeight() = True, 1, 0), 1
    GetWindowRect ComboBoxHandle, WndRect
    If ComboBoxListHandle <> NULL_PTR Then GetWindowRect ComboBoxListHandle, ListRect
    MapWindowPoints HWND_DESKTOP, ComboBoxHandle, ListRect, 2
    EditHeight = ListRect.Top
    Const SM_CYEDGE As Long = 46
    If (ListRect.Bottom - ListRect.Top) > (GetSystemMetrics(SM_CYEDGE) * 2) Then
        Height = EditHeight + (ListRect.Bottom - ListRect.Top)
    Else
        Height = EditHeight
    End If
    .Extender.Height = .ScaleY(Height, vbPixels, vbContainerSize)
    If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
    MoveWindow ComboBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    Me.Refresh
End If
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyComboBox
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

#If VBA7 Then
Public Property Get hWnd() As LongPtr
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
#Else
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
#End If
hWnd = ComboBoxHandle
End Property

#If VBA7 Then
Public Property Get hWndUserControl() As LongPtr
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
#Else
Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
#End If
hWndUserControl = UserControl.hWnd
End Property

#If VBA7 Then
Public Property Get hWndEdit() As LongPtr
Attribute hWndEdit.VB_Description = "Returns a handle to a control."
#Else
Public Property Get hWndEdit() As Long
Attribute hWndEdit.VB_Description = "Returns a handle to a control."
#End If
hWndEdit = ComboBoxEditHandle
End Property

#If VBA7 Then
Public Property Get hWndList() As LongPtr
Attribute hWndList.VB_Description = "Returns a handle to a control."
#Else
Public Property Get hWndList() As Long
Attribute hWndList.VB_Description = "Returns a handle to a control."
#End If
hWndList = ComboBoxListHandle
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
Dim OldFontHandle As LongPtr
Set PropFont = NewFont
OldFontHandle = ComboBoxFontHandle
ComboBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ComboBoxHandle <> NULL_PTR Then SendMessage ComboBoxHandle, WM_SETFONT, ComboBoxFontHandle, ByVal 1&
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As LongPtr
OldFontHandle = ComboBoxFontHandle
ComboBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ComboBoxHandle <> NULL_PTR Then SendMessage ComboBoxHandle, WM_SETFONT, ComboBoxFontHandle, ByVal 1&
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If ComboBoxHandle <> NULL_PTR And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles ComboBoxHandle
    Else
        RemoveVisualStyles ComboBoxHandle
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
If ComboBoxHandle <> NULL_PTR Then EnableWindow ComboBoxHandle, IIf(Value = True, 1, 0)
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
If ComboBoxDesignMode = False Then Call RefreshMousePointer
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
    If Value.Type = vbPicTypeIcon Or Value.Handle = NULL_PTR Then
        Set PropMouseIcon = Value
    Else
        If ComboBoxDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ComboBoxDesignMode = False Then Call RefreshMousePointer
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
If ComboBoxHandle <> NULL_PTR Then Call ComCtlsSetRightToLeft(ComboBoxHandle, dwMask)
If ComboBoxEditHandle <> NULL_PTR Then Call ComCtlsSetRightToLeft(ComboBoxEditHandle, dwMask)
If PropRightToLeft = False And ComboBoxEditHandle <> NULL_PTR Then
    Const ES_RIGHT As Long = &H2
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ComboBoxEditHandle, GWL_STYLE)
    If (dwStyle And ES_RIGHT) = ES_RIGHT Then dwStyle = dwStyle And Not ES_RIGHT
    SetWindowLong ComboBoxEditHandle, GWL_STYLE, dwStyle
End If
If ComboBoxListHandle <> NULL_PTR Then Call ComCtlsSetRightToLeft(ComboBoxListHandle, dwMask)
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
Attribute Redraw.VB_Description = "Returns/sets a value that determines whether or not the combo box redraws when changing the items. You can speed up the creation of large lists by disabling this property before adding the items."
Redraw = PropRedraw
End Property

Public Property Let Redraw(ByVal Value As Boolean)
PropRedraw = Value
If ComboBoxHandle <> NULL_PTR And ComboBoxDesignMode = False Then
    SendMessage ComboBoxHandle, WM_SETREDRAW, IIf(PropRedraw = True, 1, 0), ByVal 0&
    If PropRedraw = True Then Me.Refresh
End If
End Property

Public Property Get Style() As CboStyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines the type of control and the behavior of its list box portion."
Style = PropStyle
End Property

Public Property Let Style(ByVal Value As CboStyleConstants)
Select Case Value
    Case CboStyleDropDownCombo, CboStyleSimpleCombo, CboStyleDropDownList
        If ComboBoxDesignMode = False Then
            Err.Raise Number:=382, Description:="Style property is read-only at run time"
        Else
            PropStyle = Value
            If ComboBoxHandle <> NULL_PTR Then Call ReCreateComboBox
        End If
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "Style"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents can be edited."
Locked = PropLocked
End Property

Public Property Let Locked(ByVal Value As Boolean)
PropLocked = Value
If ComboBoxHandle <> NULL_PTR And ComboBoxEditHandle <> NULL_PTR Then SendMessage ComboBoxEditHandle, EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Locked"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "123c"
Select Case PropStyle
    Case CboStyleDropDownCombo, CboStyleSimpleCombo
        If ComboBoxHandle <> NULL_PTR And ComboBoxEditHandle <> NULL_PTR Then
            Text = String(CLng(SendMessage(ComboBoxEditHandle, WM_GETTEXTLENGTH, 0, ByVal 0&)), vbNullChar)
            SendMessage ComboBoxEditHandle, WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
        Else
            Text = PropText
        End If
    Case CboStyleDropDownList
        If ComboBoxHandle <> NULL_PTR And ComboBoxDesignMode = False Then
            Dim SelIndex As Long
            SelIndex = CLng(SendMessage(ComboBoxHandle, CB_GETCURSEL, 0, ByVal 0&))
            If Not SelIndex = CB_ERR Then Text = Me.List(SelIndex)
        Else
            Text = Ambient.DisplayName
        End If
End Select
End Property

Public Property Let Text(ByVal Value As String)
Dim Changed As Boolean
Select Case PropStyle
    Case CboStyleDropDownCombo, CboStyleSimpleCombo
        If PropMaxLength > 0 Then Value = Left$(Value, PropMaxLength)
        Changed = CBool(Me.Text <> Value)
        PropText = Value
        If ComboBoxHandle <> NULL_PTR And ComboBoxEditHandle <> NULL_PTR Then SendMessage ComboBoxEditHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
    Case CboStyleDropDownList
        If ComboBoxHandle <> NULL_PTR And ComboBoxDesignMode = False Then
            Dim Index As Long
            Index = CLng(SendMessage(ComboBoxHandle, CB_FINDSTRINGEXACT, -1, ByVal StrPtr(Value)))
            If Not Index = CB_ERR Then
                Me.ListIndex = Index
            Else
                Err.Raise Number:=383, Description:="Property is read-only"
            End If
        Else
            Exit Property
        End If
End Select
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

Public Property Get ExtendedUI() As Boolean
Attribute ExtendedUI.VB_Description = "Returns/sets a value that determines whether the default UI or the extended UI is used."
If ComboBoxHandle <> NULL_PTR And PropStyle <> CboStyleSimpleCombo Then
    ExtendedUI = CBool(SendMessage(ComboBoxHandle, CB_GETEXTENDEDUI, 0, ByVal 0&) = 1)
Else
    ExtendedUI = PropExtendedUI
End If
End Property

Public Property Let ExtendedUI(ByVal Value As Boolean)
PropExtendedUI = Value
If ComboBoxHandle <> NULL_PTR And PropStyle <> CboStyleSimpleCombo Then SendMessage ComboBoxHandle, CB_SETEXTENDEDUI, IIf(PropExtendedUI = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "ExtendedUI"
End Property

Public Property Get MaxDropDownItems() As Integer
Attribute MaxDropDownItems.VB_Description = "Returns/sets the maximum number of items to be shown in the drop-down list."
MaxDropDownItems = PropMaxDropDownItems
End Property

Public Property Let MaxDropDownItems(ByVal Value As Integer)
Select Case Value
    Case 1 To 30
        PropMaxDropDownItems = Value
        ComboBoxDropDownHeightState = False
    Case Else
        If ComboBoxDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
Call CheckDropDownHeight(True)
UserControl.PropertyChanged "MaxDropDownItems"
End Property

Public Property Get IntegralHeight() As Boolean
Attribute IntegralHeight.VB_Description = "Returns/sets a value indicating whether the control displays partial items. This property has no effect in an variable owner-drawn combo box."
IntegralHeight = PropIntegralHeight
End Property

Public Property Let IntegralHeight(ByVal Value As Boolean)
If ComboBoxDesignMode = False Then
    Err.Raise Number:=382, Description:="IntegralHeight property is read-only at run time"
Else
    PropIntegralHeight = Value
    If ComboBoxHandle <> NULL_PTR Then Call ReCreateComboBox
End If
UserControl.PropertyChanged "IntegralHeight"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
MaxLength = PropMaxLength
End Property

Public Property Let MaxLength(ByVal Value As Long)
If Value < 0 Then
    If ComboBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropMaxLength = Value
If ComboBoxHandle <> NULL_PTR Then SendMessage ComboBoxHandle, CB_LIMITTEXT, IIf(PropMaxLength = 0, 255, PropMaxLength), ByVal 0&
UserControl.PropertyChanged "MaxLength"
End Property

Public Property Get CueBanner() As String
Attribute CueBanner.VB_Description = "Returns/sets the textual cue, or tip, that is displayed to prompt the user for information. Requires comctl32.dll version 6.1 or higher."
CueBanner = PropCueBanner
End Property

Public Property Let CueBanner(ByVal Value As String)
PropCueBanner = Value
If ComboBoxHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then SendMessage ComboBoxHandle, CB_SETCUEBANNER, 0, ByVal StrPtr(PropCueBanner)
UserControl.PropertyChanged "CueBanner"
End Property

Public Property Get UseListBackColor() As Boolean
Attribute UseListBackColor.VB_Description = "Returns/sets a value which determines if the combo box control will use the list back color property."
UseListBackColor = PropUseListBackColor
End Property

Public Property Let UseListBackColor(ByVal Value As Boolean)
PropUseListBackColor = Value
Me.Refresh
UserControl.PropertyChanged "UseListBackColor"
End Property

Public Property Get UseListForeColor() As Boolean
Attribute UseListForeColor.VB_Description = "Returns/sets a value which determines if the combo box control will use the list fore color property."
UseListForeColor = PropUseListForeColor
End Property

Public Property Let UseListForeColor(ByVal Value As Boolean)
PropUseListForeColor = Value
Me.Refresh
UserControl.PropertyChanged "UseListForeColor"
End Property

Public Property Get ListBackColor() As OLE_COLOR
Attribute ListBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in the control's list portion."
ListBackColor = PropListBackColor
End Property

Public Property Let ListBackColor(ByVal Value As OLE_COLOR)
PropListBackColor = Value
If ComboBoxHandle <> NULL_PTR Then
    If ComboBoxListBackColorBrush <> NULL_PTR Then DeleteObject ComboBoxListBackColorBrush
    ComboBoxListBackColorBrush = CreateSolidBrush(WinColor(PropListBackColor))
End If
Me.Refresh
UserControl.PropertyChanged "ListBackColor"
End Property

Public Property Get ListForeColor() As OLE_COLOR
Attribute ListForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in the control's list portion."
ListForeColor = PropListForeColor
End Property

Public Property Let ListForeColor(ByVal Value As OLE_COLOR)
PropListForeColor = Value
Me.Refresh
UserControl.PropertyChanged "ListForeColor"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
Sorted = PropSorted
End Property

Public Property Let Sorted(ByVal Value As Boolean)
PropSorted = Value
If ComboBoxHandle <> NULL_PTR Then Call ReCreateComboBox
UserControl.PropertyChanged "Sorted"
End Property

Public Property Get HorizontalExtent() As Single
Attribute HorizontalExtent.VB_Description = "Returns/sets the width by which a drop-down list can be scrolled horizontally."
If ComboBoxHandle <> NULL_PTR Then
    HorizontalExtent = UserControl.ScaleX(SendMessage(ComboBoxHandle, CB_GETHORIZONTALEXTENT, 0, ByVal 0&), vbPixels, vbContainerSize)
Else
    HorizontalExtent = UserControl.ScaleX(PropHorizontalExtent, vbPixels, vbContainerSize)
End If
End Property

Public Property Let HorizontalExtent(ByVal Value As Single)
If Value < 0 Then
    If ComboBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropHorizontalExtent = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If ComboBoxHandle <> NULL_PTR Then SendMessage ComboBoxHandle, CB_SETHORIZONTALEXTENT, PropHorizontalExtent, ByVal 0&
UserControl.PropertyChanged "HorizontalExtent"
End Property

Public Property Get DisableNoScroll() As Boolean
Attribute DisableNoScroll.VB_Description = "Returns/sets a value that determines whether scroll bars are disabled instead of hided when they are not needed."
DisableNoScroll = PropDisableNoScroll
End Property

Public Property Let DisableNoScroll(ByVal Value As Boolean)
PropDisableNoScroll = Value
If ComboBoxHandle <> NULL_PTR Then Call ReCreateComboBox
UserControl.PropertyChanged "DisableNoScroll"
End Property

Public Property Get CharacterCasing() As CboCharacterCasingConstants
Attribute CharacterCasing.VB_Description = "Returns/sets a value indicating if the combo box modifies the case of characters in both the selection field and the list."
CharacterCasing = PropCharacterCasing
End Property

Public Property Let CharacterCasing(ByVal Value As CboCharacterCasingConstants)
Select Case Value
    Case CboCharacterCasingNormal, CboCharacterCasingUpper, CboCharacterCasingLower
        PropCharacterCasing = Value
    Case Else
        Err.Raise 380
End Select
If ComboBoxHandle <> NULL_PTR Then Call ReCreateComboBox
UserControl.PropertyChanged "CharacterCasing"
End Property

Public Property Get DrawMode() As CboDrawModeConstants
Attribute DrawMode.VB_Description = "Returns/sets a value indicating whether your code or the operating system will handle drawing of the elements."
DrawMode = PropDrawMode
End Property

Public Property Let DrawMode(ByVal Value As CboDrawModeConstants)
Select Case Value
    Case CboDrawModeNormal, CboDrawModeOwnerDrawFixed, CboDrawModeOwnerDrawVariable
        If ComboBoxDesignMode = False Then
            Err.Raise Number:=382, Description:="DrawMode property is read-only at run time"
        Else
            PropDrawMode = Value
            If ComboBoxHandle <> NULL_PTR Then Call ReCreateComboBox
        End If
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "DrawMode"
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
If ComboBoxHandle <> NULL_PTR And ComboBoxEditHandle <> NULL_PTR And ComboBoxDesignMode = False Then
    If GetFocus() = ComboBoxEditHandle Then Call ComCtlsSetIMEMode(ComboBoxEditHandle, ComboBoxIMCHandle, PropIMEMode)
End If
UserControl.PropertyChanged "IMEMode"
End Property

Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_Description = "Returns/sets whether the control should scroll its contents while the user moves the scroll box along the scroll bars."
ScrollTrack = PropScrollTrack
End Property

Public Property Let ScrollTrack(ByVal Value As Boolean)
PropScrollTrack = Value
UserControl.PropertyChanged "ScrollTrack"
End Property

Public Property Get AutoSelect() As Boolean
Attribute AutoSelect.VB_Description = "Returns/sets a value that determines whether or not the items can be selected automatically after an user input in the edit portion of the control."
AutoSelect = PropAutoSelect
End Property

Public Property Let AutoSelect(ByVal Value As Boolean)
PropAutoSelect = Value
UserControl.PropertyChanged "AutoSelect"
End Property

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to the combo box."
If ComboBoxHandle <> NULL_PTR Then
    If StrPtr(Item) = NULL_PTR Then Item = ""
    Dim RetVal As Long
    If IsMissing(Index) = True Then
        RetVal = CLng(SendMessage(ComboBoxHandle, CB_ADDSTRING, 0, ByVal StrPtr(Item)))
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
        RetVal = CLng(SendMessage(ComboBoxHandle, CB_INSERTSTRING, IndexLong, ByVal StrPtr(Item)))
    End If
    If Not RetVal = CB_ERR Then
        ComboBoxNewIndex = RetVal
    Else
        Err.Raise 5
    End If
    Call CheckDropDownHeight(False)
End If
End Sub

Public Sub RemoveItem(ByVal Index As Long)
Attribute RemoveItem.VB_Description = "Removes an item from the combo box."
If ComboBoxHandle <> NULL_PTR Then
    If Index >= 0 Then
        If Not SendMessage(ComboBoxHandle, CB_DELETESTRING, Index, ByVal 0&) = CB_ERR Then
            ComboBoxNewIndex = -1
        Else
            Err.Raise 5
        End If
    Else
        Err.Raise 5
    End If
    Call CheckDropDownHeight(False)
End If
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of the combo box."
If ComboBoxHandle <> NULL_PTR Then
    SendMessage ComboBoxHandle, CB_RESETCONTENT, 0, ByVal 0&
    ComboBoxNewIndex = -1
    Call CheckDropDownHeight(False)
End If
End Sub

Public Property Get ListCount() As Long
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
Attribute ListCount.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR Then ListCount = CLng(SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&))
End Property

Public Property Get List(ByVal Index As Long) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
Attribute List.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR Then
    Dim Length As Long
    Length = CLng(SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, Index, ByVal 0&))
    If Not Length = CB_ERR Then
        List = String(Length, vbNullChar)
        SendMessage ComboBoxHandle, CB_GETLBTEXT, Index, ByVal StrPtr(List)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Let List(ByVal Index As Long, ByVal Value As String)
If ComboBoxHandle <> NULL_PTR Then
    If StrPtr(Value) = NULL_PTR Then Value = ""
    If Index > -1 Then
        Dim SelIndex As Long, ItemData As LongPtr
        SelIndex = CLng(SendMessage(ComboBoxHandle, CB_GETCURSEL, 0, ByVal 0&))
        ItemData = SendMessage(ComboBoxHandle, CB_GETITEMDATA, Index, ByVal 0&)
        If Not SendMessage(ComboBoxHandle, CB_DELETESTRING, Index, ByVal 0&) = CB_ERR Then
            SendMessage ComboBoxHandle, CB_INSERTSTRING, Index, ByVal StrPtr(Value)
            SendMessage ComboBoxHandle, CB_SETCURSEL, SelIndex, ByVal 0&
            SendMessage ComboBoxHandle, CB_SETITEMDATA, Index, ByVal ItemData
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
If ComboBoxHandle <> NULL_PTR Then ListIndex = CLng(SendMessage(ComboBoxHandle, CB_GETCURSEL, 0, ByVal 0&))
End Property

Public Property Let ListIndex(ByVal Value As Long)
If ComboBoxHandle <> NULL_PTR Then
    Dim Changed As Boolean
    Changed = CBool(SendMessage(ComboBoxHandle, CB_GETCURSEL, 0, ByVal 0&) <> Value)
    If Not Value = -1 Then
        If SendMessage(ComboBoxHandle, CB_SETCURSEL, Value, ByVal 0&) = CB_ERR Then Err.Raise 380
    Else
        SendMessage ComboBoxHandle, CB_SETCURSEL, -1, ByVal 0&
    End If
    If Changed = True Then RaiseEvent Click
End If
End Property

#If VBA7 Then
Public Property Get ItemData(ByVal Index As Long) As LongPtr
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a combo box."
Attribute ItemData.VB_MemberFlags = "400"
#Else
Public Property Get ItemData(ByVal Index As Long) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a combo box."
Attribute ItemData.VB_MemberFlags = "400"
#End If
If ComboBoxHandle <> NULL_PTR Then
    If Not SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, Index, ByVal 0&) = CB_ERR Then
        ItemData = SendMessage(ComboBoxHandle, CB_GETITEMDATA, Index, ByVal 0&)
    Else
        Err.Raise 381
    End If
End If
End Property

#If VBA7 Then
Public Property Let ItemData(ByVal Index As Long, ByVal Value As LongPtr)
#Else
Public Property Let ItemData(ByVal Index As Long, ByVal Value As Long)
#End If
If ComboBoxHandle <> NULL_PTR Then
    If Not SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, Index, ByVal 0&) = CB_ERR Then
        SendMessage ComboBoxHandle, CB_SETITEMDATA, Index, ByVal Value
    Else
        Err.Raise 381
    End If
End If
End Property

Private Sub CreateComboBox()
If ComboBoxHandle <> NULL_PTR Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or CBS_AUTOHSCROLL Or WS_VSCROLL Or WS_HSCROLL
If PropRightToLeft = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR
Select Case PropStyle
    Case CboStyleDropDownCombo
        dwStyle = dwStyle Or CBS_DROPDOWN
    Case CboStyleSimpleCombo
        dwStyle = dwStyle Or CBS_SIMPLE
    Case CboStyleDropDownList
        dwStyle = dwStyle Or CBS_DROPDOWNLIST
End Select
If GetIntegralHeight() = False Then dwStyle = dwStyle Or CBS_NOINTEGRALHEIGHT
If PropSorted = True Then dwStyle = dwStyle Or CBS_SORT
If PropDisableNoScroll = True Then dwStyle = dwStyle Or CBS_DISABLENOSCROLL
Select Case PropCharacterCasing
    Case CboCharacterCasingUpper
        dwStyle = dwStyle Or CBS_UPPERCASE
    Case CboCharacterCasingLower
        dwStyle = dwStyle Or CBS_LOWERCASE
End Select
Select Case PropDrawMode
    Case CboDrawModeOwnerDrawFixed
        dwStyle = dwStyle Or CBS_OWNERDRAWFIXED Or CBS_HASSTRINGS
    Case CboDrawModeOwnerDrawVariable
        dwStyle = dwStyle Or CBS_OWNERDRAWVARIABLE Or CBS_HASSTRINGS
        If Not (dwStyle And CBS_NOINTEGRALHEIGHT) = CBS_NOINTEGRALHEIGHT Then dwStyle = dwStyle Or CBS_NOINTEGRALHEIGHT
End Select
ComboBoxHandle = CreateWindowEx(dwExStyle, StrPtr("ComboBox"), StrPtr("Combo Box"), dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
If ComboBoxHandle <> NULL_PTR Then
    Dim CBI As COMBOBOXINFO
    CBI.cbSize = LenB(CBI)
    GetComboBoxInfo ComboBoxHandle, CBI
    If PropStyle = CboStyleDropDownCombo Then
        ComboBoxEditHandle = CBI.hWndItem
        If ComboBoxEditHandle = NULL_PTR Then ComboBoxEditHandle = FindWindowEx(ComboBoxHandle, NULL_PTR, StrPtr("Edit"), NULL_PTR)
    ElseIf PropStyle = CboStyleSimpleCombo Then
        ComboBoxEditHandle = FindWindowEx(ComboBoxHandle, NULL_PTR, StrPtr("Edit"), NULL_PTR)
        If GetIntegralHeight() = False Then MoveWindow ComboBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + 2, 1
    End If
    ComboBoxListHandle = CBI.hWndList
    SendMessage ComboBoxHandle, CB_LIMITTEXT, IIf(PropMaxLength = 0, 255, PropMaxLength), ByVal 0&
    If PropStyle <> CboStyleDropDownList And ComboBoxEditHandle <> NULL_PTR Then SendMessage ComboBoxEditHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
    If PropHorizontalExtent > 0 Then SendMessage ComboBoxHandle, CB_SETHORIZONTALEXTENT, PropHorizontalExtent, ByVal 0&
    ComboBoxNewIndex = -1
    ComboBoxTopIndex = 0
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If PropRedraw = False Then Me.Redraw = False
If PropLocked = True Then Me.Locked = PropLocked
Me.ExtendedUI = PropExtendedUI
Me.MaxDropDownItems = PropMaxDropDownItems
If Not PropCueBanner = vbNullString Then Me.CueBanner = PropCueBanner
If ComboBoxDesignMode = False Then
    If ComboBoxHandle <> NULL_PTR Then
        ComboBoxInitFieldHeight = CLng(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, -1, ByVal 0&))
        If ComboBoxListBackColorBrush = NULL_PTR Then ComboBoxListBackColorBrush = CreateSolidBrush(WinColor(PropListBackColor))
        Call ComCtlsSetSubclass(ComboBoxHandle, Me, 1)
        If ComboBoxEditHandle <> NULL_PTR Then
            Call ComCtlsSetSubclass(ComboBoxEditHandle, Me, 2)
            Call ComCtlsCreateIMC(ComboBoxEditHandle, ComboBoxIMCHandle)
        End If
        If ComboBoxListHandle <> NULL_PTR Then Call ComCtlsSetSubclass(ComboBoxListHandle, Me, 3)
        Call ComCtlsSetSubclass(UserControl.hWnd, Me, 4)
    End If
Else
    If ComboBoxHandle <> NULL_PTR Then
        If ComboBoxListBackColorBrush = NULL_PTR Then ComboBoxListBackColorBrush = CreateSolidBrush(WinColor(PropListBackColor))
        Call ComCtlsSetSubclass(ComboBoxHandle, Me, 5)
    End If
    If PropStyle = CboStyleDropDownList Then
        If ComboBoxHandle <> NULL_PTR Then
            Dim Buffer As String
            Buffer = Ambient.DisplayName
            SendMessage ComboBoxHandle, CB_ADDSTRING, 0, ByVal StrPtr(Buffer)
            SendMessage ComboBoxHandle, CB_SETCURSEL, 0, ByVal 0&
        End If
    End If
End If
End Sub

Private Sub ReCreateComboBox()
If ComboBoxDesignMode = False Then
    Dim Locked As Boolean
    With Me
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Dim ListArr() As String, ItemDataArr() As LongPtr
    Dim ItemHeight As Long, ListIndex As Long, TopIndex As Long, Text As String, SelStart As Long, SelEnd As Long, DroppedWidth As Long, FieldHeight As Long, NewIndex As Long
    Dim Count As Long, i As Long, FieldHeightCustomized As Boolean
    If ComboBoxHandle <> NULL_PTR Then
        If PropDrawMode <> CboDrawModeOwnerDrawVariable Then
            ' ItemHeight does not need to be restored in an variable owner-drawn combo box as the 'MeasureItem' event will handle this.
            ItemHeight = CLng(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, 0, ByVal 0&))
        End If
        Count = CLng(SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&))
        If Count > 0 Then
            ReDim ListArr(0 To (Count - 1)) As String
            ReDim ItemDataArr(0 To (Count - 1)) ' As LongPtr
            For i = 0 To (Count - 1)
                ListArr(i) = .List(i)
                ItemDataArr(i) = SendMessage(ComboBoxHandle, CB_GETITEMDATA, i, ByVal 0&)
            Next i
        End If
        ListIndex = .ListIndex
        TopIndex = .TopIndex
        Text = .Text
        If ComboBoxEditHandle <> NULL_PTR Then SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
        DroppedWidth = CLng(SendMessage(ComboBoxHandle, CB_GETDROPPEDWIDTH, 0, ByVal 0&))
        FieldHeight = CLng(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, -1, ByVal 0&))
    End If
    NewIndex = ComboBoxNewIndex
    FieldHeightCustomized = CBool(FieldHeight <> ComboBoxInitFieldHeight)
    If FieldHeightCustomized = True Then
        Call DestroyComboBox ' This is necessary to be able to resize without any adjustments.
        With UserControl
        .Extender.Move .Extender.Left, .Extender.Top, .Extender.Width, .Extender.Height - .ScaleY((FieldHeight - ComboBoxInitFieldHeight), vbPixels, vbContainerSize)
        If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
        End With
    End If
    Call DestroyComboBox
    Call CreateComboBox
    Call UserControl_Resize
    If ComboBoxHandle <> NULL_PTR Then
        If PropDrawMode <> CboDrawModeOwnerDrawVariable Then SendMessage ComboBoxHandle, CB_SETITEMHEIGHT, 0, ByVal ItemHeight
        If Count > 0 Then
            SendMessage ComboBoxHandle, WM_SETREDRAW, 0, ByVal 0&
            For i = 0 To (Count - 1)
                SendMessage ComboBoxHandle, CB_INSERTSTRING, i, ByVal StrPtr(ListArr(i))
                SendMessage ComboBoxHandle, CB_SETITEMDATA, i, ByVal ItemDataArr(i)
            Next i
            SendMessage ComboBoxHandle, WM_SETREDRAW, 1, ByVal 0&
            Call CheckDropDownHeight(True)
        End If
        .ListIndex = ListIndex
        .TopIndex = TopIndex
        If PropStyle <> CboStyleDropDownList Then .Text = Text
        If ComboBoxEditHandle <> NULL_PTR Then SendMessage ComboBoxEditHandle, EM_SETSEL, SelStart, ByVal SelEnd
        If Not DroppedWidth = CB_ERR Then SendMessage ComboBoxHandle, CB_SETDROPPEDWIDTH, DroppedWidth, ByVal 0&
        If FieldHeightCustomized = True Then SendMessage ComboBoxHandle, CB_SETITEMHEIGHT, -1, ByVal FieldHeight
    End If
    ComboBoxNewIndex = NewIndex
    If Locked = True Then LockWindowUpdate NULL_PTR
    .Refresh
    If PropRedraw = False Then .Redraw = PropRedraw
    End With
Else
    Call DestroyComboBox
    Call CreateComboBox
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyComboBox()
If ComboBoxHandle = NULL_PTR Then Exit Sub
Call ComCtlsRemoveSubclass(ComboBoxHandle)
If ComboBoxEditHandle <> NULL_PTR Then
    Call ComCtlsRemoveSubclass(ComboBoxEditHandle)
    Call ComCtlsDestroyIMC(ComboBoxEditHandle, ComboBoxIMCHandle)
End If
If ComboBoxListHandle <> NULL_PTR Then Call ComCtlsRemoveSubclass(ComboBoxListHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow ComboBoxHandle, SW_HIDE
SetParent ComboBoxHandle, NULL_PTR
DestroyWindow ComboBoxHandle
ComboBoxHandle = NULL_PTR
ComboBoxEditHandle = NULL_PTR
ComboBoxListHandle = NULL_PTR
If ComboBoxFontHandle <> NULL_PTR Then
    DeleteObject ComboBoxFontHandle
    ComboBoxFontHandle = NULL_PTR
End If
If ComboBoxListBackColorBrush <> NULL_PTR Then
    DeleteObject ComboBoxListBackColorBrush
    ComboBoxListBackColorBrush = NULL_PTR
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
If PropRedraw = True Or ComboBoxDesignMode = True Then RedrawWindow UserControl.hWnd, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute SelStart.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR And ComboBoxEditHandle <> NULL_PTR Then SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal 0&
End Property

Public Property Let SelStart(ByVal Value As Long)
If ComboBoxHandle <> NULL_PTR And ComboBoxEditHandle <> NULL_PTR Then
    If Value >= 0 Then
        SendMessage ComboBoxEditHandle, EM_SETSEL, Value, ByVal Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR And ComboBoxEditHandle <> NULL_PTR Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    SelLength = SelEnd - SelStart
End If
End Property

Public Property Let SelLength(ByVal Value As Long)
If ComboBoxHandle <> NULL_PTR And ComboBoxEditHandle <> NULL_PTR Then
    If Value >= 0 Then
        Dim SelStart As Long
        SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal 0&
        SendMessage ComboBoxEditHandle, EM_SETSEL, SelStart, ByVal SelStart + Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR Then
    If ComboBoxEditHandle <> NULL_PTR Then
        Dim SelStart As Long, SelEnd As Long
        SendMessage ComboBoxHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
        On Error Resume Next
        SelText = Mid$(Me.Text, SelStart + 1, (SelEnd - SelStart))
        On Error GoTo 0
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Let SelText(ByVal Value As String)
If ComboBoxHandle <> NULL_PTR Then
    If ComboBoxEditHandle <> NULL_PTR Then
        If StrPtr(Value) = NULL_PTR Then Value = ""
        SendMessage ComboBoxEditHandle, EM_REPLACESEL, 0, ByVal StrPtr(Value)
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get ItemHeight(Optional ByVal Index As Long) As Single
Attribute ItemHeight.VB_Description = "Returns/sets the height of an item in the drop-down list. The optional index argument can be specified in an variable owner-drawn combo box."
Attribute ItemHeight.VB_MemberFlags = "400"
If Index < 0 Then Err.Raise 380
If ComboBoxHandle <> NULL_PTR Then
    Dim RetVal As Long
    If PropDrawMode <> CboDrawModeOwnerDrawVariable Then
        If Index = 0 Then
            RetVal = CLng(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, 0, ByVal 0&))
        Else
            RetVal = CB_ERR
        End If
    Else
        RetVal = CLng(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, Index, ByVal 0&))
    End If
    If Not RetVal = CB_ERR Then
        ItemHeight = UserControl.ScaleY(RetVal, vbPixels, vbContainerSize)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Let ItemHeight(Optional ByVal Index As Long, ByVal Value As Single)
If Value < 0 Or Index < 0 Then Err.Raise 380
If ComboBoxHandle <> NULL_PTR Then
    Dim RetVal As Long
    If PropDrawMode <> CboDrawModeOwnerDrawVariable Then
        If Index = 0 Then
            RetVal = CLng(SendMessage(ComboBoxHandle, CB_SETITEMHEIGHT, 0, ByVal CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))))
        Else
            RetVal = CB_ERR
        End If
    Else
        RetVal = CLng(SendMessage(ComboBoxHandle, CB_SETITEMHEIGHT, Index, ByVal CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))))
    End If
    If Not RetVal = CB_ERR Then
        If GetIntegralHeight() = True Then Call UserControl_Resize
        Me.Refresh
    Else
        Err.Raise 5
    End If
End If
Call CheckDropDownHeight(True)
End Property

Public Property Get FieldHeight() As Single
Attribute FieldHeight.VB_Description = "Returns/sets the height of the edit-control (or static-text) portion of the combo box."
Attribute FieldHeight.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR Then FieldHeight = UserControl.ScaleY(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, -1, ByVal 0&), vbPixels, vbContainerSize)
End Property

Public Property Let FieldHeight(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If ComboBoxHandle <> NULL_PTR Then
    If Not SendMessage(ComboBoxHandle, CB_SETITEMHEIGHT, -1, ByVal CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))) = CB_ERR Then
        Me.Refresh
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Get DroppedDown() As Boolean
Attribute DroppedDown.VB_Description = "Returns/sets a value that determines whether the drop-down list is dropped down or not."
Attribute DroppedDown.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR Then DroppedDown = CBool(SendMessage(ComboBoxHandle, CB_GETDROPPEDSTATE, 0, ByVal 0&) <> 0)
End Property

Public Property Let DroppedDown(ByVal Value As Boolean)
If ComboBoxHandle <> NULL_PTR Then SendMessage ComboBoxHandle, CB_SHOWDROPDOWN, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Property Get DropDownWidth() As Single
Attribute DropDownWidth.VB_Description = "Returns/sets the width of the drop-down list. This property is not supported in a simple combo box."
Attribute DropDownWidth.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR Then
    Dim RetVal As Long
    RetVal = CLng(SendMessage(ComboBoxHandle, CB_GETDROPPEDWIDTH, 0, ByVal 0&))
    If Not RetVal = CB_ERR Then
        DropDownWidth = UserControl.ScaleX(RetVal, vbPixels, vbContainerSize)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Let DropDownWidth(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If ComboBoxHandle <> NULL_PTR Then
    If SendMessage(ComboBoxHandle, CB_SETDROPPEDWIDTH, CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels)), ByVal 0&) = CB_ERR Then Err.Raise 5
End If
End Property

Public Property Get DropDownHeight() As Single
Attribute DropDownHeight.VB_Description = "Returns/sets the height of the drop-down list. Setting this property resets the integral height property to false. Also the max drop-down items property gets not meaningful anymore. This property is not supported in a simple combo box."
Attribute DropDownHeight.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR Then
    If PropStyle <> CboStyleSimpleCombo Then
        Dim ListRect As RECT
        If ComboBoxListHandle <> NULL_PTR Then GetWindowRect ComboBoxListHandle, ListRect
        DropDownHeight = UserControl.ScaleY((ListRect.Bottom - ListRect.Top), vbPixels, vbContainerSize)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Let DropDownHeight(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If ComboBoxHandle <> NULL_PTR Then
    If PropStyle <> CboStyleSimpleCombo Then
        Dim LngValue As Long
        LngValue = CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
        If LngValue > 0 Then
            If GetIntegralHeight() = True Then
                PropIntegralHeight = False
                Call ReCreateComboBox
            End If
            ComboBoxDropDownHeightState = True
            MoveWindow ComboBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + LngValue, 1
        Else
            Err.Raise 380
        End If
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Get NewIndex() As Long
Attribute NewIndex.VB_Description = "Returns the index of the item most recently added to a control."
Attribute NewIndex.VB_MemberFlags = "400"
NewIndex = ComboBoxNewIndex
End Property

Public Property Get TopIndex() As Long
Attribute TopIndex.VB_Description = "Returns/sets which item in a control is displayed in the topmost position."
Attribute TopIndex.VB_MemberFlags = "400"
If ComboBoxHandle <> NULL_PTR Then TopIndex = CLng(SendMessage(ComboBoxHandle, CB_GETTOPINDEX, 0, ByVal 0&))
End Property

Public Property Let TopIndex(ByVal Value As Long)
If ComboBoxHandle <> NULL_PTR Then
    If Value >= 0 Then
        If SendMessage(ComboBoxHandle, CB_SETTOPINDEX, Value, ByVal 0&) = CB_ERR Then Err.Raise 380
    Else
        Err.Raise 380
    End If
End If
End Property

Public Function FindItem(ByVal Text As String, Optional ByVal Index As Long = -1, Optional ByVal Partial As Boolean) As Long
Attribute FindItem.VB_Description = "Finds an item in the combo box and returns the index of that item."
If ComboBoxHandle <> NULL_PTR Then
    If Not SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, Index, ByVal 0&) = CB_ERR Or Index = -1 Then
        If Partial = True Then
            FindItem = CLng(SendMessage(ComboBoxHandle, CB_FINDSTRING, Index, ByVal StrPtr(Text)))
        Else
            FindItem = CLng(SendMessage(ComboBoxHandle, CB_FINDSTRINGEXACT, Index, ByVal StrPtr(Text)))
        End If
    Else
        Err.Raise 381
    End If
End If
End Function

Public Function GetIdealHorizontalExtent() As Single
Attribute GetIdealHorizontalExtent.VB_Description = "Gets the ideal value for the horizontal extent property."
If ComboBoxHandle <> NULL_PTR And ComboBoxListHandle <> NULL_PTR Then
    Dim Count As Long
    Count = CLng(SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&))
    If Count > 0 Then
        Dim RC(0 To 1) As RECT, CX As Long, ScrollWidth As Long, hDC As LongPtr, i As Long, Length As Long, Text As String, Size As SIZEAPI
        GetWindowRect ComboBoxListHandle, RC(0)
        GetClientRect ComboBoxListHandle, RC(1)
        If (GetWindowLong(ComboBoxListHandle, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL Then
            Const SM_CXVSCROLL As Long = 2
            ScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
        End If
        hDC = GetDC(ComboBoxHandle)
        Dim hFontOld As LongPtr
        hFontOld = SelectObject(hDC, ComboBoxFontHandle)
        For i = 0 To Count - 1
            Length = CLng(SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, i, ByVal 0&))
            If Not Length = CB_ERR Then
                Text = String(Length, vbNullChar)
                SendMessage ComboBoxHandle, CB_GETLBTEXT, i, ByVal StrPtr(Text)
                GetTextExtentPoint32 hDC, ByVal StrPtr(Text), Length, Size
                If (Size.CX - ScrollWidth) > CX Then CX = (Size.CX - ScrollWidth)
            End If
        Next i
        If hFontOld <> NULL_PTR Then SelectObject hDC, hFontOld
        ReleaseDC ComboBoxHandle, hDC
        If CX > 0 Then GetIdealHorizontalExtent = UserControl.ScaleX(CX + ((RC(0).Right - RC(0).Left) - (RC(1).Right - RC(1).Left)), vbPixels, vbContainerSize)
    End If
End If
End Function

Public Function SelectItem(ByVal Text As String, Optional ByVal Index As Long = -1) As Long
Attribute SelectItem.VB_Description = "Searches for an item that begins with the characters in a specified string. If a matching item is found, the item is selected. The search is not case sensitive."
If ComboBoxHandle <> NULL_PTR Then
    If Not SendMessage(ComboBoxHandle, CB_GETLBTEXTLEN, Index, ByVal 0&) = CB_ERR Or Index = -1 Then
        Dim OldIndex As Long
        OldIndex = CLng(SendMessage(ComboBoxHandle, CB_GETCURSEL, 0, ByVal 0&))
        SelectItem = CLng(SendMessage(ComboBoxHandle, CB_SELECTSTRING, Index, ByVal StrPtr(Text)))
        If SelectItem <> OldIndex And Not SelectItem = CB_ERR Then RaiseEvent Click
    Else
        Err.Raise 381
    End If
End If
End Function

Private Sub CheckDropDownHeight(ByVal Calculate As Boolean)
Static LastCount As Long, ItemHeight As Long
If ComboBoxHandle <> NULL_PTR And ComboBoxDropDownHeightState = False Then
    Dim Count As Long, Height As Long
    Count = CLng(SendMessage(ComboBoxHandle, CB_GETCOUNT, 0, ByVal 0&))
    If PropDrawMode <> CboDrawModeOwnerDrawVariable Then
        Select Case Count
            Case 0
                Count = 1
            Case Is > PropMaxDropDownItems
                Count = PropMaxDropDownItems
        End Select
        If Calculate = False Then
            If Count = LastCount Then Exit Sub
        Else
            ItemHeight = CLng(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, 0, ByVal 0&))
        End If
        Height = (ItemHeight * Count)
    Else
        If Calculate = False Then Exit Sub
        If Count > 0 Then
            Dim TopIndex As Long, i As Long, RetVal As Long
            TopIndex = CLng(SendMessage(ComboBoxHandle, CB_GETTOPINDEX, 0, ByVal 0&))
            If Not TopIndex = CB_ERR Then
                For i = TopIndex To (TopIndex + (PropMaxDropDownItems - 1))
                    RetVal = CLng(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, i, ByVal 0&))
                    If RetVal = CB_ERR Then Exit For
                    Height = Height + RetVal
                Next i
            End If
        Else
            Height = CLng(SendMessage(ComboBoxHandle, CB_GETITEMHEIGHT, -1, ByVal 0&))
        End If
    End If
    If PropStyle <> CboStyleSimpleCombo Then
        MoveWindow ComboBoxHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + Height + 2, 1
        If GetIntegralHeight() = True And ComCtlsSupportLevel() >= 1 Then SendMessage ComboBoxHandle, CB_SETMINVISIBLE, PropMaxDropDownItems, ByVal 0&
    Else
        RedrawWindow ComboBoxHandle, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
    End If
    LastCount = Count
End If
End Sub

Private Sub CheckTopIndex()
Dim TopIndex As Long
If ComboBoxHandle <> NULL_PTR Then TopIndex = CLng(SendMessage(ComboBoxHandle, CB_GETTOPINDEX, 0, ByVal 0&))
If TopIndex <> ComboBoxTopIndex Then
    ComboBoxTopIndex = TopIndex
    RaiseEvent Scroll
End If
End Sub

Private Sub CheckAutoSelect()
If PropAutoSelect = True Then
    Select Case PropStyle
        Case CboStyleDropDownCombo, CboStyleSimpleCombo
            Dim Index As Long
            If ComboBoxHandle <> NULL_PTR Then
                Index = CLng(SendMessage(ComboBoxHandle, CB_FINDSTRINGEXACT, -1, ByVal StrPtr(Me.Text)))
                If Not Index = CB_ERR Then
                    Me.ListIndex = Index
                    Me.SelStart = Len(Me.Text)
                End If
            End If
    End Select
End If
End Sub

Private Function GetIntegralHeight() As Boolean
If PropDrawMode <> CboDrawModeOwnerDrawVariable Then
    GetIntegralHeight = PropIntegralHeight
Else
    GetIntegralHeight = False
End If
End Function

#If VBA7 Then
Private Function ISubclass_Message(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
#End If
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcEdit(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcList(hWnd, wMsg, wParam, lParam)
    Case 4
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Case 5
        ISubclass_Message = WindowProcControlDesignMode(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd And (wParam <> ComboBoxEditHandle Or ComboBoxEditHandle = NULL_PTR) Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_LBUTTONDOWN
        If ComboBoxEditHandle = NULL_PTR Then
            If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
        Else
            Select Case GetFocus()
                Case hWnd, ComboBoxEditHandle
                Case Else
                    UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
            End Select
        End If
    Case WM_SETCURSOR
        If LoWord(CLng(lParam)) = HTCLIENT Then
            If MousePointerID(PropMousePointer) <> 0 Then
                SetCursor LoadCursor(NULL_PTR, MousePointerID(PropMousePointer))
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
    Case WM_CTLCOLORLISTBOX
        WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If PropUseListBackColor = True Or PropUseListForeColor = True Then SetBkMode wParam, 1
        If PropUseListForeColor = True Then SetTextColor wParam, WinColor(PropListForeColor)
        If PropUseListBackColor = True And ComboBoxListBackColorBrush <> NULL_PTR Then WindowProcControl = ComboBoxListBackColorBrush
        Exit Function
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        If PropStyle = CboStyleDropDownList Then
            Dim KeyCode As Integer
            KeyCode = CLng(wParam) And &HFF&
            If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
                If wMsg = WM_KEYDOWN Then
                    RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
                ElseIf wMsg = WM_KEYUP Then
                    RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
                End If
                ComboBoxCharCodeCache = ComCtlsPeekCharCode(hWnd)
            ElseIf wMsg = WM_SYSKEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_SYSKEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            wParam = KeyCode
        End If
    Case WM_CHAR
        If PropStyle = CboStyleDropDownList Then
            Dim KeyChar As Integer
            If ComboBoxCharCodeCache <> 0 Then
                KeyChar = CUIntToInt(ComboBoxCharCodeCache And &HFFFF&)
                ComboBoxCharCodeCache = 0
            Else
                KeyChar = CUIntToInt(CLng(wParam) And &HFFFF&)
            End If
            RaiseEvent KeyPress(KeyChar)
            wParam = CIntToUInt(KeyChar)
        End If
    Case WM_UNICHAR
        If PropStyle = CboStyleDropDownList Then
            If wParam = UNICODE_NOCHAR Then
                WindowProcControl = 1
            Else
                Dim UTF16 As String
                UTF16 = UTF32CodePoint_To_UTF16(CLng(wParam))
                If Len(UTF16) = 1 Then
                    SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
                ElseIf Len(UTF16) = 2 Then
                    SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                    SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
                End If
                WindowProcControl = 0
            End If
            Exit Function
        End If
    Case WM_IME_CHAR
        If PropStyle = CboStyleDropDownList Then
            SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
            Exit Function
        End If
    Case WM_CONTEXTMENU
        If wParam = ComboBoxHandle Then
            Dim P1 As POINTAPI, Handled As Boolean
            P1.X = Get_X_lParam(lParam)
            P1.Y = Get_Y_lParam(lParam)
            If P1.X = -1 And P1.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(Handled, -1, -1)
            Else
                ScreenToClient ComboBoxHandle, P1
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P1.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P1.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
    Case WM_SIZE
        If ComboBoxResizeFrozen = False Then
            Dim WndRect As RECT
            GetWindowRect hWnd, WndRect
            With UserControl
            If (WndRect.Bottom - WndRect.Top) <> .ScaleHeight Or (WndRect.Right - WndRect.Left) <> .ScaleWidth Then
                ComboBoxResizeFrozen = True
                .Extender.Move .Extender.Left, .Extender.Top, .ScaleX((WndRect.Right - WndRect.Left), vbPixels, vbContainerSize), .ScaleY((WndRect.Bottom - WndRect.Top), vbPixels, vbContainerSize)
                If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
                ComboBoxResizeFrozen = False
            End If
            End With
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
                If (ComboBoxMouseOver(0) = False And PropMouseTrack = True) Or (ComboBoxMouseOver(2) = False And PropMouseTrack = True) Then
                    If ComboBoxMouseOver(0) = False And PropMouseTrack = True Then ComboBoxMouseOver(0) = True
                    If ComboBoxMouseOver(2) = False And PropMouseTrack = True Then
                        ComboBoxMouseOver(2) = True
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
        End Select
    Case WM_MOUSELEAVE
        ComboBoxMouseOver(0) = False
        If ComboBoxMouseOver(2) = True Then
            Dim Pos As Long, P2 As POINTAPI, XY As Currency
            Pos = GetMessagePos()
            P2.X = Get_X_lParam(Pos)
            P2.Y = Get_Y_lParam(Pos)
            CopyMemory ByVal VarPtr(XY), ByVal VarPtr(P2), 8
            If WindowFromPoint(XY) <> ComboBoxEditHandle Or ComboBoxEditHandle = NULL_PTR Then
                ComboBoxMouseOver(2) = False
                RaiseEvent MouseLeave
            End If
        End If
    Case CB_SETTOPINDEX
        Call CheckTopIndex
End Select
End Function

Private Function WindowProcEdit(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd And wParam <> ComboBoxHandle Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_SETCURSOR
        If LoWord(CLng(lParam)) = HTCLIENT Then
            If PropOLEDragMode = vbOLEDragAutomatic Then
                Dim P1 As POINTAPI
                Dim CharPos As Long, CaretPos As Long
                Dim SelStart As Long, SelEnd As Long
                GetCursorPos P1
                ScreenToClient ComboBoxEditHandle, P1
                CharPos = LoWord(CLng(SendMessage(ComboBoxEditHandle, EM_CHARFROMPOS, 0, ByVal MakeDWord(P1.X, P1.Y))))
                CaretPos = CLng(SendMessage(ComboBoxEditHandle, EM_POSFROMCHAR, CharPos, ByVal 0&))
                SendMessage ComboBoxEditHandle, EM_GETSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
                ComboBoxAutoDragInSel = CBool(CharPos >= SelStart And CharPos <= SelEnd And CaretPos > -1 And (SelEnd - SelStart) > 0)
                If ComboBoxAutoDragInSel = True Then
                    ComboBoxAutoDragSelStart = SelStart
                    ComboBoxAutoDragSelEnd = SelEnd
                    SetCursor LoadCursor(NULL_PTR, MousePointerID(vbArrow))
                    WindowProcEdit = 1
                    Exit Function
                End If
            Else
                ComboBoxAutoDragInSel = False
            End If
        End If
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = CLng(wParam) And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            ComboBoxCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If ComboBoxCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(ComboBoxCharCodeCache And &HFFFF&)
            ComboBoxCharCodeCache = 0
        Else
            KeyChar = CUIntToInt(CLng(wParam) And &HFFFF&)
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
            UTF16 = UTF32CodePoint_To_UTF16(CLng(wParam))
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcEdit = 0
        End If
        Exit Function
    Case WM_INPUTLANGCHANGE
        Call ComCtlsSetIMEMode(hWnd, ComboBoxIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call ComCtlsSetIMEMode(hWnd, ComboBoxIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_LBUTTONDOWN
        If PropOLEDragMode = vbOLEDragAutomatic And ComboBoxAutoDragInSel = True Then
            Dim P2 As POINTAPI, XY1 As Currency
            P2.X = Get_X_lParam(lParam)
            P2.Y = Get_Y_lParam(lParam)
            ClientToScreen ComboBoxEditHandle, P2
            CopyMemory ByVal VarPtr(XY1), ByVal VarPtr(P2), 8
            If DragDetect(ComboBoxEditHandle, XY1) <> 0 Then
                Me.OLEDrag
                WindowProcEdit = 0
            Else
                WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
                ReleaseCapture
            End If
            Exit Function
        Else
            Select Case GetFocus()
                Case hWnd, ComboBoxHandle
                Case Else
                    UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
            End Select
        End If
    Case WM_CONTEXTMENU
        If wParam = hWnd Then
            Dim P3 As POINTAPI, Handled As Boolean
            P3.X = Get_X_lParam(lParam)
            P3.Y = Get_Y_lParam(lParam)
            If P3.X = -1 And P3.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(Handled, -1, -1)
            Else
                ScreenToClient ComboBoxHandle, P3
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P3.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P3.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
End Select
WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim P4 As POINTAPI
        P4.X = Get_X_lParam(lParam)
        P4.Y = Get_Y_lParam(lParam)
        If ComboBoxHandle <> NULL_PTR Then MapWindowPoints hWnd, ComboBoxHandle, P4, 1
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(P4.X, vbPixels, vbTwips)
        Y = UserControl.ScaleY(P4.Y, vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MOUSEMOVE
                If (ComboBoxMouseOver(1) = False And PropMouseTrack = True) Or (ComboBoxMouseOver(2) = False And PropMouseTrack = True) Then
                    If ComboBoxMouseOver(1) = False And PropMouseTrack = True Then ComboBoxMouseOver(1) = True
                    If ComboBoxMouseOver(2) = False And PropMouseTrack = True Then
                        ComboBoxMouseOver(2) = True
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
        End Select
    Case WM_MOUSELEAVE
        ComboBoxMouseOver(1) = False
        If ComboBoxMouseOver(2) = True Then
            Dim Pos As Long, P5 As POINTAPI, XY2 As Currency
            Pos = GetMessagePos()
            P5.X = Get_X_lParam(Pos)
            P5.Y = Get_Y_lParam(Pos)
            CopyMemory ByVal VarPtr(XY2), ByVal VarPtr(P5), 8
            If WindowFromPoint(XY2) <> ComboBoxHandle Or ComboBoxHandle = NULL_PTR Then
                ComboBoxMouseOver(2) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function

Private Function WindowProcList(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_CHAR
        If PropLocked = True Then Exit Function
    Case WM_KEYDOWN, WM_KEYUP
        If PropLocked = True Then
            Dim KeyCode As Integer
            KeyCode = CLng(wParam) And &HFF&
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
                    Exit Function
            End Select
        End If
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP, WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        If PropLocked = True Then
            Dim P As POINTAPI, XY As Currency
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            ClientToScreen hWnd, P
            CopyMemory ByVal VarPtr(XY), ByVal VarPtr(P), 8
            If Not LBItemFromPt(hWnd, XY, 0) = LB_ERR Then Exit Function
        End If
    Case WM_VSCROLL
        Select Case LoWord(CLng(wParam))
            Case SB_THUMBPOSITION, SB_THUMBTRACK
                ' HiWord carries only 16 bits of scroll box position data.
                ' Below workaround will circumvent the 16-bit barrier by using the 32-bit GetScrollInfo function.
                Dim dwStyle As Long
                dwStyle = GetWindowLong(ComboBoxListHandle, GWL_STYLE)
                If lParam = 0 And (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                    Dim SCI As SCROLLINFO, PrevPos As Long
                    SCI.cbSize = LenB(SCI)
                    SCI.fMask = SIF_POS Or SIF_TRACKPOS
                    GetScrollInfo ComboBoxListHandle, SB_VERT, SCI
                    PrevPos = SCI.nPos
                    Select Case LoWord(CLng(wParam))
                        Case SB_THUMBPOSITION
                            SCI.nPos = SCI.nTrackPos
                        Case SB_THUMBTRACK
                            If PropScrollTrack = True Then SCI.nPos = SCI.nTrackPos
                    End Select
                    If PrevPos <> SCI.nPos Then
                        ' SetScrollInfo function not needed as CB_SETTOPINDEX itself will do the scrolling.
                        SendMessage ComboBoxHandle, CB_SETTOPINDEX, SCI.nPos, ByVal 0&
                    End If
                    WindowProcList = 0
                    Exit Function
                End If
        End Select
End Select
WindowProcList = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_MOUSEMOVE
        If (GetMouseStateFromParam(wParam) And vbLeftButton) = vbLeftButton Then Call CheckTopIndex
    Case WM_MOUSEWHEEL, WM_VSCROLL, LB_SETTOPINDEX
        Call CheckTopIndex
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_COMMAND
        Select Case HiWord(CLng(wParam))
            Case CBN_SELCHANGE
                Dim SelIndex As Long
                SelIndex = CLng(SendMessage(lParam, CB_GETCURSEL, 0, ByVal 0&))
                If Not SelIndex = CB_ERR Then
                    If PropStyle <> CboStyleDropDownList And ComboBoxEditHandle <> NULL_PTR Then SendMessage ComboBoxEditHandle, WM_SETTEXT, 0, ByVal StrPtr(Me.List(SelIndex))
                    Call CheckTopIndex
                    RaiseEvent Click
                End If
            Case CBN_DBLCLK
                RaiseEvent DblClick
            Case CBN_EDITCHANGE
                UserControl.PropertyChanged "Text"
                On Error Resume Next
                UserControl.Extender.DataChanged = True
                On Error GoTo 0
                Call CheckAutoSelect
                RaiseEvent Change
            Case CBN_DROPDOWN
                If PropStyle <> CboStyleDropDownList And ComboBoxEditHandle <> NULL_PTR Then
                    If GetCursor() = NULL_PTR Then
                        ' The mouse cursor can be hidden when showing the drop-down list upon a change event.
                        ' Reason is that the edit control hides the cursor and a following mouse move will show it again.
                        ' However, the drop-down list will set a mouse capture and thus the cursor keeps hidden.
                        ' Solution is to refresh the cursor by sending a WM_SETCURSOR.
                        Call RefreshMousePointer(lParam)
                    End If
                End If
                If PropDrawMode = CboDrawModeOwnerDrawVariable Then Call CheckDropDownHeight(True)
                RaiseEvent DropDown
            Case CBN_CLOSEUP
                RaiseEvent CloseUp
        End Select
    Case WM_MEASUREITEM
        If PropDrawMode = CboDrawModeOwnerDrawVariable Then
            Dim MIS As MEASUREITEMSTRUCT
            CopyMemory MIS, ByVal lParam, LenB(MIS)
            If MIS.CtlType = ODT_COMBOBOX And MIS.ItemID > -1 Then
                With MIS
                RaiseEvent ItemMeasure(.ItemID, .ItemHeight)
                End With
                CopyMemory ByVal lParam, MIS, LenB(MIS)
                WindowProcUserControl = 1
                Exit Function
            End If
        End If
    Case WM_DRAWITEM
        If PropDrawMode <> CboDrawModeNormal Then
            Dim DIS As DRAWITEMSTRUCT
            CopyMemory DIS, ByVal lParam, LenB(DIS)
            If DIS.CtlType = ODT_COMBOBOX And DIS.hWndItem = ComboBoxHandle And DIS.ItemID > -1 Then
                With DIS
                #If Win64 Then
                Dim hDC32 As Long
                CopyMemory ByVal VarPtr(hDC32), ByVal VarPtr(.hDC), 4
                RaiseEvent ItemDraw(.ItemID, .ItemAction, .ItemState, hDC32, .RCItem.Left, .RCItem.Top, .RCItem.Right, .RCItem.Bottom)
                #Else
                RaiseEvent ItemDraw(.ItemID, .ItemAction, .ItemState, .hDC, .RCItem.Left, .RCItem.Top, .RCItem.Right, .RCItem.Bottom)
                #End If
                End With
                WindowProcUserControl = 1
                Exit Function
            End If
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI ComboBoxHandle
End Function

Private Function WindowProcControlDesignMode(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_CTLCOLORLISTBOX
        WindowProcControlDesignMode = WindowProcControl(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
WindowProcControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
