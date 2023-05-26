VERSION 5.00
Begin VB.UserControl ImageCombo 
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   PropertyPages   =   "ImageCombo.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ImageCombo.ctx":003B
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ImageCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private ImcStyleDropDownCombo, ImcStyleSimpleCombo, ImcStyleDropDownList
Private ImcEndEditReasonLostFocus, ImcEndEditReasonReturn, ImcEndEditReasonDropDown
Private ImcEllipsisFormatNone, ImcEllipsisFormatEnd
#End If
Public Enum ImcStyleConstants
ImcStyleDropDownCombo = 0
ImcStyleSimpleCombo = 1
ImcStyleDropDownList = 2
End Enum
Private Const CBENF_KILLFOCUS As Long = &H1
Private Const CBENF_RETURN As Long = &H2
Private Const CBENF_ESCAPE As Long = &H3
Private Const CBENF_DROPDOWN As Long = &H4
Public Enum ImcEndEditReasonConstants
ImcEndEditReasonLostFocus = CBENF_KILLFOCUS
ImcEndEditReasonReturn = CBENF_RETURN
ImcEndEditReasonEscape = CBENF_ESCAPE
ImcEndEditReasonDropDown = CBENF_DROPDOWN
End Enum
Public Enum ImcEllipsisFormatConstants
ImcEllipsisFormatNone = 0
ImcEllipsisFormatEnd = 1
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
hWndCombo As Long
hWndItem As Long
hWndList As Long
End Type
Private Type COMBOBOXEXITEM
Mask As Long
iItem As Long
pszText As Long
cchTextMax As Long
iImage As Long
iSelectedImage As Long
iOverlay As Long
iIndent As Long
lParam As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMCOMBOBOXEX
hdr As NMHDR
CeItem As COMBOBOXEXITEM
End Type
Private Const CBEMAXSTRLEN As Long = 260
Private Type NMCBEDRAGBEGIN
hdr As NMHDR
iItem As Long
szText(0 To ((CBEMAXSTRLEN * 2) - 1)) As Byte
End Type
Private Type NMCBEENDEDIT
hdr As NMHDR
fChanged As Long
iNewSelection As Long
szText(0 To ((CBEMAXSTRLEN * 2) - 1)) As Byte
iWhy As Long
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
Public Event DropDown()
Attribute DropDown.VB_Description = "Occurs when the drop-down list is about to drop down."
Public Event CloseUp()
Attribute CloseUp.VB_Description = "Occurs when the drop-down list has been closed."
Public Event ItemDrag(ByVal Item As ImcComboItem, ByVal Button As Integer)
Attribute ItemDrag.VB_Description = "Occurs when a combo item initiate a drag-and-drop operation."
Public Event BeginEdit()
Attribute BeginEdit.VB_Description = "Occurs when the user activates the drop-down list or clicks in the edit field."
Public Event EndEdit(ByVal Changed As Boolean, ByVal NewIndex As Long, ByVal NewText As String, ByVal Reason As ImcEndEditReasonConstants)
Attribute EndEdit.VB_Description = "Occurs when the user has concluded an operation within the edit field or has selected an item from the drop-down list."
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
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hWndCombo As Long, ByRef CBI As COMBOBOXINFO) As Long
Private Declare Function LBItemFromPt Lib "comctl32" (ByVal hLB As Long, ByVal PX As Long, ByVal PY As Long, ByVal bAutoScroll As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, ByRef CX As Long, ByRef CY As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByRef lpScrollInfo As SCROLLINFO) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Const ICC_USEREX_CLASSES As Long = &H200
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const GWL_ID As Long = (-12)
Private Const GWL_STYLE As Long = (-16)
Private Const CF_UNICODETEXT As Long = 13
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const SW_HIDE As Long = &H0
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_NOTIFY As Long = &H4E
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
Private Const WM_CHARTOITEM As Long = &H2F
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
Private Const WM_VSCROLL As Long = &H115
Private Const SB_VERT As Long = 1
Private Const SB_THUMBPOSITION As Long = 4, SB_THUMBTRACK As Long = 5
Private Const SIF_POS As Long = &H4
Private Const SIF_TRACKPOS As Long = &H10
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const EM_SETREADONLY As Long = &HCF
Private Const EM_SETSEL As Long = &HB1
Private Const EM_REPLACESEL As Long = &HC2
Private Const LB_ERR As Long = (-1)
Private Const LB_SETTOPINDEX As Long = &H197
Private Const CB_ERR As Long = (-1)
Private Const CB_LIMITTEXT As Long = &H141
Private Const CB_DELETESTRING As Long = &H144
Private Const CB_GETCOUNT As Long = &H146
Private Const CB_GETCURSEL As Long = &H147
Private Const CB_SETCURSEL As Long = &H14E
Private Const CB_GETTOPINDEX As Long = &H15B
Private Const CB_SETTOPINDEX As Long = &H15C
Private Const CB_GETDROPPEDWIDTH As Long = &H15F
Private Const CB_SETDROPPEDWIDTH As Long = &H160
Private Const CB_GETLBTEXT As Long = &H148
Private Const CB_GETLBTEXTLEN As Long = &H149
Private Const CB_GETEDITSEL As Long = &H140
Private Const CB_SETEDITSEL As Long = &H142
Private Const CB_RESETCONTENT As Long = &H14B
Private Const CB_SETITEMHEIGHT As Long = &H153
Private Const CB_GETITEMHEIGHT As Long = &H154
Private Const CB_GETDROPPEDSTATE As Long = &H157
Private Const CB_GETCOMBOBOXINFO As Long = &H164 ' Unsupported on W2K
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const CB_SETEXTENDEDUI As Long = &H155
Private Const CB_GETEXTENDEDUI As Long = &H156
Private Const CBS_AUTOHSCROLL As Long = &H40
Private Const CBS_SIMPLE As Long = &H1
Private Const CBS_DROPDOWN As Long = &H2
Private Const CBS_DROPDOWNLIST As Long = &H3
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETUNICODEFORMAT As Long = (CCM_FIRST + 5)
Private Const WM_USER As Long = &H400
Private Const UM_SETFOCUS As Long = (WM_USER + 444)
Private Const UM_BUTTONDOWN As Long = (WM_USER + 500)
Private Const CBEM_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT
Private Const CBEM_INSERTITEMA As Long = (WM_USER + 1)
Private Const CBEM_INSERTITEMW As Long = (WM_USER + 11)
Private Const CBEM_INSERTITEM As Long = CBEM_INSERTITEMW
Private Const CBEM_SETIMAGELIST As Long = (WM_USER + 2)
Private Const CBEM_GETIMAGELIST As Long = (WM_USER + 3)
Private Const CBEM_GETITEMA As Long = (WM_USER + 4)
Private Const CBEM_GETITEMW As Long = (WM_USER + 13)
Private Const CBEM_GETITEM As Long = CBEM_GETITEMW
Private Const CBEM_SETITEMA As Long = (WM_USER + 5)
Private Const CBEM_SETITEMW As Long = (WM_USER + 12)
Private Const CBEM_SETITEM As Long = CBEM_SETITEMW
Private Const CBEM_DELETEITEM As Long = CB_DELETESTRING
Private Const CBEM_GETCOMBOCONTROL As Long = (WM_USER + 6)
Private Const CBEM_GETEDITCONTROL As Long = (WM_USER + 7)
Private Const CBEM_SETEXTENDEDSTYLE As Long = (WM_USER + 8)
Private Const CBEM_GETEXTENDEDSTYLE As Long = (WM_USER + 9)
Private Const CBEM_HASEDITCHANGED As Long = (WM_USER + 10)
Private Const CBEIF_TEXT As Long = &H1
Private Const CBEIF_IMAGE As Long = &H2
Private Const CBEIF_SELECTEDIMAGE As Long = &H4
Private Const CBEIF_OVERLAY As Long = &H8
Private Const CBEIF_INDENT As Long = &H10
Private Const CBEIF_LPARAM As Long = &H20
Private Const CBES_EX_NOEDITIMAGE As Long = &H1
Private Const CBES_EX_NOEDITIMAGEINDENT As Long = &H2
Private Const CBES_EX_PATHWORDBREAKPROC As Long = &H4
Private Const CBES_EX_NOSIZELIMIT As Long = &H8
Private Const CBES_EX_CASESENSITIVE As Long = &H10
Private Const CBES_EX_TEXTENDELLIPSIS As Long = &H20
Private Const I_IMAGECALLBACK As Long = (-1)
Private Const CBEN_FIRST As Long = (-800)
Private Const CBEN_GETDISPINFOA As Long = (CBEN_FIRST - 0)
Private Const CBEN_GETDISPINFOW As Long = (CBEN_FIRST - 7)
Private Const CBEN_GETDISPINFO As Long = CBEN_GETDISPINFOW
Private Const CBEN_INSERTITEM As Long = (CBEN_FIRST - 1)
Private Const CBEN_DELETEITEM As Long = (CBEN_FIRST - 2)
Private Const CBEN_BEGINEDIT As Long = (CBEN_FIRST - 4)
Private Const CBEN_ENDEDITA As Long = (CBEN_FIRST - 5)
Private Const CBEN_ENDEDITW As Long = (CBEN_FIRST - 6)
Private Const CBEN_ENDEDIT As Long = CBEN_ENDEDITW
Private Const CBEN_DRAGBEGINA As Long = (CBEN_FIRST - 8)
Private Const CBEN_DRAGBEGINW As Long = (CBEN_FIRST - 9)
Private Const CBEN_DRAGBEGIN As Long = CBEN_DRAGBEGINW
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
Private ImageComboHandle As Long
Private ImageComboComboHandle As Long, ImageComboEditHandle As Long, ImageComboListHandle As Long
Private ImageComboFontHandle As Long
Private ImageComboIMCHandle As Long
Private ImageComboCharCodeCache As Long
Private ImageComboMouseOver(0 To 2) As Boolean
Private ImageComboDesignMode As Boolean
Private ImageComboTopIndex As Long
Private ImageComboDragIndexBuffer As Long, ImageComboDragIndex As Long
Private ImageComboImageListObjectPointer As Long
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private DispIDImageList As Long, ImageListArray() As String
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropComboItems As ImcComboItems
Private PropVisualStyles As Boolean
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropImageListName As String, PropImageListInit As Boolean
Private PropStyle As ImcStyleConstants
Private PropLocked As Boolean
Private PropText As String
Private PropIndentation As Long
Private PropExtendedUI As Boolean
Private PropMaxDropDownItems As Integer
Private PropShowImages As Boolean
Private PropMaxLength As Long
Private PropIMEMode As CCIMEModeConstants
Private PropEllipsisFormat As ImcEllipsisFormatConstants
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
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyTab, vbKeyReturn, vbKeyEscape
            If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
                If SendMessage(ImageComboHandle, CB_GETDROPPEDSTATE, 0, ByVal 0&) = 0 Or PropStyle = ImcStyleSimpleCombo Then
                    If IsInputKey = False Then Exit Sub
                Else
                    If PropStyle = ImcStyleDropDownCombo Or PropStyle = ImcStyleDropDownList Then SendMessage ImageComboHandle, CB_SHOWDROPDOWN, 0, ByVal 0&
                End If
            ElseIf KeyCode = vbKeyTab Then
                If SendMessage(ImageComboHandle, CB_GETDROPPEDSTATE, 0, ByVal 0&) = 1 Then SendMessage ImageComboHandle, CB_SHOWDROPDOWN, 0, ByVal 0&
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
ElseIf DispID = DispIDImageList Then
    DisplayName = PropImageListName
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
    Handled = True
ElseIf DispID = DispIDImageList Then
    On Error GoTo CATCH_EXCEPTION
    Call ComCtlsIPPBSetPredefinedStringsImageList(StringsOut(), CookiesOut(), UserControl.ParentControls, ImageListArray())
    On Error GoTo 0
    Handled = True
End If
Exit Sub
CATCH_EXCEPTION:
Handled = False
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Then
    Value = Cookie
    Handled = True
ElseIf DispID = DispIDImageList Then
    If Cookie < UBound(ImageListArray()) Then Value = ImageListArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_USEREX_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim ImageListArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
ImageComboDesignMode = Not Ambient.UserMode
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
PropImageListName = "(None)"
PropStyle = ImcStyleDropDownCombo
PropLocked = False
PropText = Ambient.DisplayName
PropIndentation = 0
PropExtendedUI = False
PropMaxDropDownItems = 9
PropShowImages = True
PropMaxLength = 0
PropIMEMode = CCIMEModeNoControl
PropEllipsisFormat = ImcEllipsisFormatNone
PropScrollTrack = True
Call CreateImageCombo
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
ImageComboDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
PropOLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = .ReadProperty("ImageList", "(None)")
PropStyle = .ReadProperty("Style", ImcStyleDropDownCombo)
PropLocked = .ReadProperty("Locked", False)
PropText = VarToStr(.ReadProperty("Text", vbNullString))
PropIndentation = .ReadProperty("Indentation", 0)
PropExtendedUI = .ReadProperty("ExtendedUI", False)
PropMaxDropDownItems = .ReadProperty("MaxDropDownItems", 9)
PropShowImages = .ReadProperty("ShowImages", True)
PropMaxLength = .ReadProperty("MaxLength", 0)
PropIMEMode = .ReadProperty("IMEMode", CCIMEModeNoControl)
PropEllipsisFormat = .ReadProperty("EllipsisFormat", ImcEllipsisFormatNone)
PropScrollTrack = .ReadProperty("ScrollTrack", True)
End With
Call CreateImageCombo
If Not PropImageListName = "(None)" Then TimerImageList.Enabled = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "ImageList", PropImageListName, "(None)"
.WriteProperty "Style", PropStyle, ImcStyleDropDownCombo
.WriteProperty "Locked", PropLocked, False
.WriteProperty "Text", StrToVar(PropText), vbNullString
.WriteProperty "Indentation", PropIndentation, 0
.WriteProperty "ExtendedUI", PropExtendedUI, False
.WriteProperty "MaxDropDownItems", PropMaxDropDownItems, 9
.WriteProperty "ShowImages", PropShowImages, True
.WriteProperty "MaxLength", PropMaxLength, 0
.WriteProperty "IMEMode", PropIMEMode, CCIMEModeNoControl
.WriteProperty "EllipsisFormat", PropEllipsisFormat, ImcEllipsisFormatNone
.WriteProperty "ScrollTrack", PropScrollTrack, True
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
ImageComboDragIndex = 0
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
If ImageComboDragIndex > 0 Then
    If PropOLEDragMode = vbOLEDragAutomatic Then
        Dim Text As String
        Text = Me.FComboItemText(ImageComboDragIndex)
        Data.SetData StrToVar(Text & vbNullChar), CF_UNICODETEXT
        Data.SetData Text, vbCFText
        AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    End If
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then ImageComboDragIndex = 0
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
If ImageComboDragIndex > 0 Then Exit Sub
If ImageComboDragIndexBuffer > 0 Then ImageComboDragIndex = ImageComboDragIndexBuffer
UserControl.OLEDrag
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If ImageComboDesignMode = True And PropertyName = "DisplayName" And PropStyle = ImcStyleDropDownList Then
    If ImageComboHandle <> 0 Then
        If SendMessage(ImageComboHandle, CB_GETCOUNT, 0, ByVal 0&) > 0 Then
            Me.FComboItemsClear
            Me.FComboItemsAdd 1, Ambient.DisplayName
            Me.FComboItemSelected(1) = True
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
If ImageComboHandle = 0 Then InProc = False: Exit Sub
Dim WndRect As RECT
If PropStyle <> ImcStyleSimpleCombo Then
    MoveWindow ImageComboHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    GetWindowRect ImageComboHandle, WndRect
    If (WndRect.Bottom - WndRect.Top) <> .ScaleHeight Or (WndRect.Right - WndRect.Left) <> .ScaleWidth Then
        .Extender.Height = .ScaleY((WndRect.Bottom - WndRect.Top), vbPixels, vbContainerSize)
        If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
    End If
    ' Call SetDropDownHeight(True) is not needed as 'ImageComboComboHandle' is not touched.
Else
    Dim ListRect As RECT, EditHeight As Long, ItemHeight As Long
    Dim Height As Long, Temp As Long, Count As Long
    MoveWindow ImageComboHandle, 0, 0, .ScaleWidth, 100, 1
    GetWindowRect ImageComboHandle, WndRect
    If ImageComboListHandle <> 0 Then GetWindowRect ImageComboListHandle, ListRect
    EditHeight = (WndRect.Bottom - WndRect.Top) - (ListRect.Bottom - ListRect.Top)
    MoveWindow ImageComboHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    GetWindowRect ImageComboHandle, WndRect
    ItemHeight = SendMessage(ImageComboHandle, CB_GETITEMHEIGHT, 0, ByVal 0&)
    Temp = (WndRect.Bottom - WndRect.Top) - EditHeight
    If Temp > 0 Then
        Do
            Temp = Temp - ItemHeight
            If Temp > 0 Then
                Count = Count + 1
            Else
                Exit Do
            End If
        Loop
    End If
    If Count > 0 Then
        Const SM_CYEDGE As Long = 46
        Height = EditHeight + (ItemHeight * Count) + (GetSystemMetrics(SM_CYEDGE) * 2)
    Else
        Height = EditHeight
    End If
    .Extender.Height = .ScaleY(Height, vbPixels, vbContainerSize)
    If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
End If
MoveWindow ImageComboHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyImageCombo
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropImageListInit = False Then
    Me.ImageList = PropImageListName
    PropImageListInit = True
End If
TimerImageList.Enabled = False
End Sub

Public Property Get ControlsEnum() As VBRUN.ParentControls
Attribute ControlsEnum.VB_MemberFlags = "40"
Set ControlsEnum = UserControl.ParentControls
End Property

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
hWnd = ImageComboHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndCombo() As Long
Attribute hWndCombo.VB_Description = "Returns a handle to a control."
hWndCombo = ImageComboComboHandle
End Property

Public Property Get hWndEdit() As Long
Attribute hWndEdit.VB_Description = "Returns a handle to a control."
hWndEdit = ImageComboEditHandle
End Property

Public Property Get hWndList() As Long
Attribute hWndList.VB_Description = "Returns a handle to a control."
hWndList = ImageComboListHandle
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
OldFontHandle = ImageComboFontHandle
ImageComboFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, WM_SETFONT, ImageComboFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = ImageComboFontHandle
ImageComboFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, WM_SETFONT, ImageComboFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If ImageComboHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles ImageComboComboHandle
    Else
        RemoveVisualStyles ImageComboComboHandle
    End If
    Me.Refresh
    SetWindowPos ImageComboHandle, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
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
If ImageComboHandle <> 0 Then EnableWindow ImageComboHandle, IIf(Value = True, 1, 0)
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
If ImageComboDesignMode = False Then Call RefreshMousePointer
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
        If ImageComboDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ImageComboDesignMode = False Then Call RefreshMousePointer
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
If ImageComboDesignMode = False Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
    dwMask = 0
End If
If PropRightToLeft = True Then
    If PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = WS_EX_RTLREADING Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR
End If
If ImageComboHandle <> 0 Then Call ComCtlsSetRightToLeft(ImageComboHandle, dwMask)
If ImageComboComboHandle <> 0 Then Call ComCtlsSetRightToLeft(ImageComboComboHandle, dwMask)
If ImageComboEditHandle <> 0 Then Call ComCtlsSetRightToLeft(ImageComboEditHandle, dwMask)
If (PropRightToLeft = False Or PropRightToLeftLayout = False) And ImageComboEditHandle <> 0 <> 0 Then
    Const ES_RIGHT As Long = &H2
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ImageComboEditHandle, GWL_STYLE)
    If (dwStyle And ES_RIGHT) = ES_RIGHT Then dwStyle = dwStyle And Not ES_RIGHT
    SetWindowLong ImageComboEditHandle, GWL_STYLE, dwStyle
End If
If ImageComboListHandle <> 0 Then Call ComCtlsSetRightToLeft(ImageComboListHandle, dwMask)
If ImageComboHandle <> 0 Then SetWindowPos ImageComboHandle, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
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

Public Property Get ImageList() As Variant
Attribute ImageList.VB_Description = "Returns/sets the image list control to be used."
If ImageComboDesignMode = False Then
    If PropImageListInit = False And ImageComboImageListObjectPointer = 0 Then
        If Not PropImageListName = "(None)" Then Me.ImageList = PropImageListName
        PropImageListInit = True
    End If
    Set ImageList = PropImageListControl
Else
    ImageList = PropImageListName
End If
End Property

Public Property Set ImageList(ByVal Value As Variant)
Me.ImageList = Value
End Property

Public Property Let ImageList(ByVal Value As Variant)
If ImageComboHandle <> 0 Then
    Dim Success As Boolean, Handle As Long
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> 0)
        End If
        If Success = True Then
            SendMessage ImageComboHandle, CBEM_SETIMAGELIST, 0, ByVal Handle
            ImageComboImageListObjectPointer = ObjPtr(Value)
            PropImageListName = ProperControlName(Value)
        End If
    ElseIf VarType(Value) = vbString Then
        Dim ControlEnum As Object, CompareName As String
        For Each ControlEnum In UserControl.ParentControls
            If TypeName(ControlEnum) = "ImageList" Then
                CompareName = ProperControlName(ControlEnum)
                If CompareName = Value And Not CompareName = vbNullString Then
                    Err.Clear
                    Handle = ControlEnum.hImageList
                    Success = CBool(Err.Number = 0 And Handle <> 0)
                    If Success = True Then
                        SendMessage ImageComboHandle, CBEM_SETIMAGELIST, 0, ByVal Handle
                        If ImageComboDesignMode = False Then ImageComboImageListObjectPointer = ObjPtr(ControlEnum)
                        PropImageListName = Value
                        Exit For
                    ElseIf ImageComboDesignMode = True Then
                        PropImageListName = Value
                        Success = True
                        Exit For
                    End If
                End If
            End If
        Next ControlEnum
    End If
    On Error GoTo 0
    If Success = False Then
        SendMessage ImageComboHandle, CBEM_SETIMAGELIST, 0, ByVal 0&
        ImageComboImageListObjectPointer = 0
        PropImageListName = "(None)"
    ElseIf Handle = 0 Then
        SendMessage ImageComboHandle, CBEM_SETIMAGELIST, 0, ByVal 0&
    End If
    On Error Resume Next
    Call UserControl_Resize
    On Error GoTo 0
    Call SetDropDownHeight(True)
    SetWindowPos ImageComboHandle, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
End If
UserControl.PropertyChanged "ImageList"
End Property

Public Property Get Style() As ImcStyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines the type of control and the behavior of its list box portion."
Style = PropStyle
End Property

Public Property Let Style(ByVal Value As ImcStyleConstants)
Select Case Value
    Case ImcStyleDropDownCombo, ImcStyleSimpleCombo, ImcStyleDropDownList
        If ImageComboDesignMode = False Then
            Err.Raise Number:=382, Description:="Style property is read-only at run time"
        Else
            PropStyle = Value
            If ImageComboHandle <> 0 Then
                Call DestroyImageCombo
                Call CreateImageCombo
                Call UserControl_Resize
                If Not PropImageListName = "(None)" Then Me.ImageList = PropImageListName
            End If
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
If ImageComboHandle <> 0 And ImageComboEditHandle <> 0 Then SendMessage ImageComboEditHandle, EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Locked"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "123c"
Select Case PropStyle
    Case ImcStyleDropDownCombo
        If ImageComboHandle <> 0 Then
            Dim CBEI As COMBOBOXEXITEM, Buffer As String
            Buffer = String(CBEMAXSTRLEN, vbNullChar) & vbNullChar
            With CBEI
            .Mask = CBEIF_TEXT
            .iItem = -1
            .pszText = StrPtr(Buffer)
            .cchTextMax = Len(Buffer)
            SendMessage ImageComboHandle, CBEM_GETITEM, 0, ByVal VarPtr(CBEI)
            Text = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
            End With
        Else
            Text = PropText
        End If
    Case ImcStyleSimpleCombo
        If ImageComboHandle <> 0 And ImageComboEditHandle <> 0 Then
            Text = String(SendMessage(ImageComboEditHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
            SendMessage ImageComboEditHandle, WM_GETTEXT, Len(Text) + 1, ByVal StrPtr(Text)
        Else
            Text = PropText
        End If
    Case ImcStyleDropDownList
        If ImageComboComboHandle <> 0 And ImageComboDesignMode = False Then
            Dim iItem As Long
            iItem = SendMessage(ImageComboComboHandle, CB_GETCURSEL, 0, ByVal 0&)
            If Not iItem = CB_ERR Then
                Text = String(SendMessage(ImageComboComboHandle, CB_GETLBTEXTLEN, iItem, ByVal 0&), vbNullChar)
                SendMessage ImageComboComboHandle, CB_GETLBTEXT, iItem, ByVal StrPtr(Text)
            End If
        Else
            Text = Ambient.DisplayName
        End If
End Select
End Property

Public Property Let Text(ByVal Value As String)
Select Case PropStyle
    Case ImcStyleDropDownCombo
        If PropMaxLength > 0 Then Value = Left$(Value, PropMaxLength)
        PropText = Value
        If ImageComboHandle <> 0 Then
            Dim CBEI As COMBOBOXEXITEM
            With CBEI
            .Mask = CBEIF_TEXT
            .iItem = -1
            .pszText = StrPtr(PropText)
            .cchTextMax = Len(PropText)
            End With
            SendMessage ImageComboHandle, CBEM_SETITEM, 0, ByVal VarPtr(CBEI)
        End If
    Case ImcStyleSimpleCombo
        If PropMaxLength > 0 Then Value = Left$(Value, PropMaxLength)
        PropText = Value
        If ImageComboHandle <> 0 And ImageComboEditHandle <> 0 Then SendMessage ImageComboEditHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
    Case ImcStyleDropDownList
        If ImageComboDesignMode = False Then
            Dim Item As ImcComboItem
            Set Item = Me.FindItem(Value)
            If Not Item Is Nothing Then
                Me.SelectedItem = Item
            Else
                Err.Raise Number:=383, Description:="Property is read-only"
            End If
        Else
            Exit Property
        End If
End Select
UserControl.PropertyChanged "Text"
End Property

Public Property Get Default() As String
Attribute Default.VB_UserMemId = 0
Attribute Default.VB_MemberFlags = "40"
Default = Me.Text
End Property

Public Property Let Default(ByVal Value As String)
Me.Text = Value
End Property

Public Property Get Indentation() As Long
Attribute Indentation.VB_Description = "Returns/sets default indentation in icon width for newly added combo items."
Indentation = PropIndentation
End Property

Public Property Let Indentation(ByVal Value As Long)
If Value < 0 Then
    If ImageComboDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropIndentation = Value
UserControl.PropertyChanged "Indentation"
End Property

Public Property Get ExtendedUI() As Boolean
Attribute ExtendedUI.VB_Description = "Returns/sets a value that determines whether the default UI or the extended UI is used."
If ImageComboHandle <> 0 And PropStyle <> ImcStyleSimpleCombo Then
    ExtendedUI = CBool(SendMessage(ImageComboHandle, CB_GETEXTENDEDUI, 0, ByVal 0&) = 1)
Else
    ExtendedUI = PropExtendedUI
End If
End Property

Public Property Let ExtendedUI(ByVal Value As Boolean)
PropExtendedUI = Value
If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, CB_SETEXTENDEDUI, IIf(PropExtendedUI = True, 1, 0), ByVal 0&
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
    Case Else
        If ImageComboDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
Call SetDropDownHeight(True)
UserControl.PropertyChanged "MaxDropDownItems"
End Property

Public Property Get ShowImages() As Boolean
Attribute ShowImages.VB_Description = "Returns/sets a value that determines whether the edit box and the drop-down list will display item images or not."
ShowImages = PropShowImages
End Property

Public Property Let ShowImages(ByVal Value As Boolean)
PropShowImages = Value
If ImageComboHandle <> 0 Then
    If PropShowImages = True Then
        SendMessage ImageComboHandle, CBEM_SETEXTENDEDSTYLE, CBES_EX_NOEDITIMAGE And CBES_EX_NOEDITIMAGEINDENT, ByVal 0&
    Else
        SendMessage ImageComboHandle, CBEM_SETEXTENDEDSTYLE, CBES_EX_NOEDITIMAGEINDENT, ByVal CBES_EX_NOEDITIMAGEINDENT
    End If
    SetWindowPos ImageComboHandle, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
End If
UserControl.PropertyChanged "ShowImages"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
MaxLength = PropMaxLength
End Property

Public Property Let MaxLength(ByVal Value As Long)
If Value < 0 Then
    If ImageComboDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropMaxLength = Value
If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, CB_LIMITTEXT, IIf(PropMaxLength = 0, CBEMAXSTRLEN - 1, PropMaxLength), ByVal 0&
UserControl.PropertyChanged "MaxLength"
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
If ImageComboHandle <> 0 And ImageComboEditHandle <> 0 And ImageComboDesignMode = False Then
    If GetFocus() = ImageComboEditHandle Then Call ComCtlsSetIMEMode(ImageComboEditHandle, ImageComboIMCHandle, PropIMEMode)
End If
UserControl.PropertyChanged "IMEMode"
End Property

Public Property Get EllipsisFormat() As ImcEllipsisFormatConstants
Attribute EllipsisFormat.VB_Description = "Returns/sets a value indicating if and where the ellipsis character is appended, denoting that the text extends beyond the edge of the control. Requires comctl32.dll version 6.1 or higher."
EllipsisFormat = PropEllipsisFormat
End Property

Public Property Let EllipsisFormat(ByVal Value As ImcEllipsisFormatConstants)
Select Case Value
    Case ImcEllipsisFormatNone, ImcEllipsisFormatEnd
        PropEllipsisFormat = Value
    Case Else
        Err.Raise 380
End Select
If ImageComboHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Select Case PropEllipsisFormat
        Case ImcEllipsisFormatNone
            SendMessage ImageComboHandle, CBEM_SETEXTENDEDSTYLE, CBES_EX_TEXTENDELLIPSIS, ByVal 0&
        Case ImcEllipsisFormatEnd
            SendMessage ImageComboHandle, CBEM_SETEXTENDEDSTYLE, CBES_EX_TEXTENDELLIPSIS, ByVal CBES_EX_TEXTENDELLIPSIS
    End Select
    Me.Refresh
End If
UserControl.PropertyChanged "EllipsisFormat"
End Property

Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_Description = "Returns/sets whether the control should scroll its contents while the user moves the scroll box along the scroll bars."
ScrollTrack = PropScrollTrack
End Property

Public Property Let ScrollTrack(ByVal Value As Boolean)
PropScrollTrack = Value
UserControl.PropertyChanged "ScrollTrack"
End Property

Public Property Get ComboItems() As ImcComboItems
Attribute ComboItems.VB_Description = "Returns a reference to a collection of the combo item objects."
If PropComboItems Is Nothing Then
    Set PropComboItems = New ImcComboItems
    PropComboItems.FInit Me
End If
Set ComboItems = PropComboItems
End Property

Friend Sub FComboItemsAdd(ByVal Index As Long, Optional ByVal Text As String, Optional ByVal ImageIndex As Long, Optional ByVal SelImageIndex As Long, Optional ByVal Indentation As Variant)
Dim CBEI As COMBOBOXEXITEM
With CBEI
.Mask = CBEIF_TEXT Or CBEIF_IMAGE Or CBEIF_SELECTEDIMAGE Or CBEIF_LPARAM Or CBEIF_INDENT
.iItem = Index - 1
.pszText = StrPtr(Text)
.cchTextMax = Len(Text)
.iImage = ImageIndex - 1
.iSelectedImage = SelImageIndex - 1
.lParam = 0
If IsMissing(Indentation) = True Then
    .iIndent = PropIndentation
Else
    Select Case VarType(Indentation)
        Case vbLong, vbInteger, vbByte
            If Indentation >= 0 Then
                .iIndent = Indentation
            Else
                Err.Raise 380
            End If
        Case vbDouble, vbSingle
            If CLng(Indentation) >= 0 Then
                .iIndent = CLng(Indentation)
            Else
                Err.Raise 380
            End If
        Case vbEmpty
            .iIndent = 0
        Case Else
            Err.Raise 13
    End Select
End If
End With
If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, CBEM_INSERTITEM, 0, ByVal VarPtr(CBEI)
Call SetDropDownHeight(False)
End Sub

Friend Sub FComboItemsRemove(ByVal Index As Long)
If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, CBEM_DELETEITEM, Index - 1, ByVal 0&
Call SetDropDownHeight(False)
End Sub

Friend Sub FComboItemsClear()
If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, CB_RESETCONTENT, 0, ByVal 0&
Call SetDropDownHeight(False)
End Sub

Friend Property Get FComboItemText(ByVal Index As Long) As String
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM, Buffer As String
    Buffer = String(CBEMAXSTRLEN, vbNullChar) & vbNullChar
    With CBEI
    .Mask = CBEIF_TEXT
    .iItem = Index - 1
    .pszText = StrPtr(Buffer)
    .cchTextMax = Len(Buffer)
    SendMessage ImageComboHandle, CBEM_GETITEM, 0, ByVal VarPtr(CBEI)
    FComboItemText = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End With
End If
End Property

Friend Property Let FComboItemText(ByVal Index As Long, ByVal Value As String)
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_TEXT
    .iItem = Index - 1
    .pszText = StrPtr(Value)
    .cchTextMax = Len(Value)
    End With
    SendMessage ImageComboHandle, CBEM_SETITEM, 0, ByVal VarPtr(CBEI)
End If
End Property

Friend Property Get FComboItemImage(ByVal Index As Long) As Long
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_IMAGE
    .iItem = Index - 1
    SendMessage ImageComboHandle, CBEM_GETITEM, 0, ByVal VarPtr(CBEI)
    FComboItemImage = .iImage + 1
    End With
End If
End Property

Friend Property Let FComboItemImage(ByVal Index As Long, ByVal Value As Long)
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_IMAGE
    .iItem = Index - 1
    .iImage = Value - 1
    End With
    SendMessage ImageComboHandle, CBEM_SETITEM, 0, ByVal VarPtr(CBEI)
End If
End Property

Friend Property Get FComboItemSelImage(ByVal Index As Long) As Long
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_SELECTEDIMAGE
    .iItem = Index - 1
    SendMessage ImageComboHandle, CBEM_GETITEM, 0, ByVal VarPtr(CBEI)
    FComboItemSelImage = .iSelectedImage + 1
    End With
End If
End Property

Friend Property Let FComboItemSelImage(ByVal Index As Long, ByVal Value As Long)
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_SELECTEDIMAGE
    .iItem = Index - 1
    .iSelectedImage = Value - 1
    End With
    SendMessage ImageComboHandle, CBEM_SETITEM, 0, ByVal VarPtr(CBEI)
End If
End Property

Friend Property Get FComboItemIndentation(ByVal Index As Long) As Long
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_INDENT
    .iItem = Index - 1
    SendMessage ImageComboHandle, CBEM_GETITEM, 0, ByVal VarPtr(CBEI)
    FComboItemIndentation = .iIndent
    End With
End If
End Property

Friend Property Let FComboItemIndentation(ByVal Index As Long, ByVal Value As Long)
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_INDENT
    .iItem = Index - 1
    .iIndent = Value
    End With
    SendMessage ImageComboHandle, CBEM_SETITEM, 0, ByVal VarPtr(CBEI)
End If
End Property

Friend Property Get FComboItemSelected(ByVal Index As Long) As Boolean
If ImageComboHandle <> 0 Then FComboItemSelected = CBool(SendMessage(ImageComboHandle, CB_GETCURSEL, 0, ByVal 0&) = Index - 1)
End Property

Friend Property Let FComboItemSelected(ByVal Index As Long, ByVal Value As Boolean)
If ImageComboHandle <> 0 Then
    Dim Changed As Boolean
    Changed = CBool(SendMessage(ImageComboHandle, CB_GETCURSEL, 0, ByVal 0&) <> (Index - 1))
    If Value = True Then
        SendMessage ImageComboHandle, CB_SETCURSEL, Index - 1, ByVal 0&
        If PropStyle = ImcStyleDropDownCombo Then Me.Text = Me.ComboItems(Index).Text
    Else
        If SendMessage(ImageComboHandle, CB_GETCURSEL, 0, ByVal 0&) = Index - 1 Then SendMessage ImageComboHandle, CB_SETCURSEL, -1, ByVal 0&
    End If
    SetWindowPos ImageComboHandle, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
    If Changed = True Then RaiseEvent Click
End If
End Property

Friend Property Get FComboItemData(ByVal Index As Long) As Long
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_LPARAM
    .iItem = Index - 1
    SendMessage ImageComboHandle, CBEM_GETITEM, 0, ByVal VarPtr(CBEI)
    FComboItemData = .lParam
    End With
End If
End Property

Friend Property Let FComboItemData(ByVal Index As Long, ByVal Value As Long)
If ImageComboHandle <> 0 Then
    Dim CBEI As COMBOBOXEXITEM
    With CBEI
    .Mask = CBEIF_LPARAM
    .iItem = Index - 1
    .lParam = Value
    End With
    SendMessage ImageComboHandle, CBEM_SETITEM, 0, ByVal VarPtr(CBEI)
End If
End Property

Private Sub CreateImageCombo()
If ImageComboHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or CBS_AUTOHSCROLL
If PropRightToLeft = True Then
    If PropRightToLeftLayout = True Then
        dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    Else
        dwExStyle = dwExStyle Or WS_EX_RTLREADING Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR
    End If
End If
Select Case PropStyle
    Case ImcStyleDropDownCombo
        dwStyle = dwStyle Or CBS_DROPDOWN
    Case ImcStyleSimpleCombo
        dwStyle = dwStyle Or CBS_SIMPLE
    Case ImcStyleDropDownList
        dwStyle = dwStyle Or CBS_DROPDOWNLIST
End Select
ImageComboHandle = CreateWindowEx(dwExStyle, StrPtr("ComboBoxEx32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If ImageComboHandle <> 0 Then
    SendMessage ImageComboHandle, CBEM_SETUNICODEFORMAT, 1, ByVal 0&
    ImageComboComboHandle = SendMessage(ImageComboHandle, CBEM_GETCOMBOCONTROL, 0, ByVal 0&)
    If ImageComboComboHandle <> 0 Then
        Dim CBI As COMBOBOXINFO
        CBI.cbSize = LenB(CBI)
        GetComboBoxInfo ImageComboComboHandle, CBI
    End If
    If PropStyle = ImcStyleDropDownCombo Then
        ImageComboEditHandle = SendMessage(ImageComboHandle, CBEM_GETEDITCONTROL, 0, ByVal 0&)
        If ImageComboEditHandle = 0 Then ImageComboEditHandle = FindWindowEx(ImageComboComboHandle, 0, StrPtr("Edit"), 0)
    ElseIf PropStyle = ImcStyleSimpleCombo Then
        ImageComboEditHandle = FindWindowEx(ImageComboComboHandle, 0, StrPtr("Edit"), 0)
    End If
    ImageComboListHandle = CBI.hWndList
    SendMessage ImageComboHandle, CB_LIMITTEXT, IIf(PropMaxLength = 0, CBEMAXSTRLEN - 1, PropMaxLength), ByVal 0&
    If PropStyle = ImcStyleDropDownCombo Then
        Dim CBEI As COMBOBOXEXITEM
        With CBEI
        .Mask = CBEIF_TEXT
        .iItem = -1
        .pszText = StrPtr(PropText)
        .cchTextMax = Len(PropText)
        End With
        SendMessage ImageComboHandle, CBEM_SETITEM, 0, ByVal VarPtr(CBEI)
    ElseIf PropStyle = ImcStyleSimpleCombo Then
        If ImageComboEditHandle <> 0 Then SendMessage ImageComboEditHandle, WM_SETTEXT, 0, ByVal StrPtr(PropText)
    End If
    ImageComboTopIndex = 0
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
If PropLocked = True Then Me.Locked = PropLocked
Me.Indentation = PropIndentation
Me.ExtendedUI = PropExtendedUI
Me.MaxDropDownItems = PropMaxDropDownItems
If PropShowImages = False Then Me.ShowImages = PropShowImages
If PropEllipsisFormat <> ImcEllipsisFormatNone Then Me.EllipsisFormat = PropEllipsisFormat
If ImageComboDesignMode = False Then
    If ImageComboHandle <> 0 Then
        Call ComCtlsSetSubclass(ImageComboHandle, Me, 1)
        If ImageComboComboHandle <> 0 Then Call ComCtlsSetSubclass(ImageComboComboHandle, Me, 2)
        If ImageComboEditHandle <> 0 Then
            Call ComCtlsSetSubclass(ImageComboEditHandle, Me, 3)
            Call ComCtlsCreateIMC(ImageComboEditHandle, ImageComboIMCHandle)
        End If
        If ImageComboListHandle <> 0 Then Call ComCtlsSetSubclass(ImageComboListHandle, Me, 4)
    End If
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 5)
Else
    If PropStyle = ImcStyleDropDownList Then
        Me.FComboItemsAdd 1, Ambient.DisplayName
        If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, CB_SETCURSEL, 0, ByVal 0&
    End If
End If
End Sub

Private Sub DestroyImageCombo()
If ImageComboHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(ImageComboHandle)
If ImageComboComboHandle <> 0 Then Call ComCtlsRemoveSubclass(ImageComboComboHandle)
If ImageComboEditHandle <> 0 Then
    Call ComCtlsRemoveSubclass(ImageComboEditHandle)
    Call ComCtlsDestroyIMC(ImageComboEditHandle, ImageComboIMCHandle)
End If
If ImageComboListHandle <> 0 Then Call ComCtlsRemoveSubclass(ImageComboListHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow ImageComboHandle, SW_HIDE
SetParent ImageComboHandle, 0
DestroyWindow ImageComboHandle
ImageComboHandle = 0
ImageComboComboHandle = 0
ImageComboEditHandle = 0
If ImageComboFontHandle <> 0 Then
    DeleteObject ImageComboFontHandle
    ImageComboFontHandle = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute SelStart.VB_MemberFlags = "400"
If ImageComboHandle <> 0 And ImageComboEditHandle <> 0 Then SendMessage ImageComboHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal 0&
End Property

Public Property Let SelStart(ByVal Value As Long)
If ImageComboHandle <> 0 And ImageComboEditHandle <> 0 Then
    If Value >= 0 Then
        SendMessage ImageComboEditHandle, EM_SETSEL, Value, ByVal Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
If ImageComboHandle <> 0 And ImageComboEditHandle <> 0 Then
    Dim SelStart As Long, SelEnd As Long
    SendMessage ImageComboHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
    SelLength = SelEnd - SelStart
End If
End Property

Public Property Let SelLength(ByVal Value As Long)
If ImageComboHandle <> 0 And ImageComboEditHandle <> 0 Then
    If Value >= 0 Then
        Dim SelStart As Long
        SendMessage ImageComboHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal 0&
        SendMessage ImageComboEditHandle, EM_SETSEL, SelStart, ByVal SelStart + Value
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
If ImageComboHandle <> 0 Then
    If ImageComboEditHandle <> 0 Then
        Dim SelStart As Long, SelEnd As Long
        SendMessage ImageComboHandle, CB_GETEDITSEL, VarPtr(SelStart), ByVal VarPtr(SelEnd)
        On Error Resume Next
        SelText = Mid(Me.Text, SelStart + 1, (SelEnd - SelStart))
        On Error GoTo 0
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Let SelText(ByVal Value As String)
If ImageComboHandle <> 0 Then
    If ImageComboEditHandle <> 0 Then
        If StrPtr(Value) = 0 Then Value = ""
        SendMessage ImageComboEditHandle, EM_REPLACESEL, 0, ByVal StrPtr(Value)
    Else
        Err.Raise 380
    End If
End If
End Property

Public Function GetItemHeight() As Single
Attribute GetItemHeight.VB_Description = "Determines the height of a combo item in the drop-down list."
If ImageComboHandle <> 0 Then
    Dim ItemHeight As Long, ImageHeight As Long, hImageList As Long
    ItemHeight = SendMessage(ImageComboHandle, CB_GETITEMHEIGHT, 0, ByVal 0&)
    hImageList = SendMessage(ImageComboHandle, CBEM_GETIMAGELIST, 0, ByVal 0&)
    If hImageList <> 0 Then ImageList_GetIconSize hImageList, 0, ImageHeight
    If ImageHeight > ItemHeight Then ItemHeight = ImageHeight
    GetItemHeight = UserControl.ScaleY(ItemHeight, vbPixels, vbContainerSize)
End If
End Function

Public Property Get TopItem() As ImcComboItem
Attribute TopItem.VB_Description = "Returns/sets a reference to the topmost visible combo item."
Attribute TopItem.VB_MemberFlags = "400"
If ImageComboComboHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ImageComboComboHandle, CB_GETTOPINDEX, 0, ByVal 0&)
    If Not iItem = CB_ERR Then Set TopItem = Me.ComboItems(iItem + 1)
End If
End Property

Public Property Let TopItem(ByVal Value As ImcComboItem)
Set Me.TopItem = Value
End Property

Public Property Set TopItem(ByVal Value As ImcComboItem)
If ImageComboComboHandle <> 0 Then
    If Not Value Is Nothing Then
        If SendMessage(ImageComboComboHandle, CB_SETTOPINDEX, Value.Index - 1, ByVal 0&) = CB_ERR Then Err.Raise 380
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelectedItem() As ImcComboItem
Attribute SelectedItem.VB_Description = "Returns/sets a reference to the currently selected combo item."
Attribute SelectedItem.VB_MemberFlags = "400"
If ImageComboHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ImageComboHandle, CB_GETCURSEL, 0, ByVal 0&)
    If Not iItem = CB_ERR Then Set SelectedItem = Me.ComboItems(iItem + 1)
End If
End Property

Public Property Let SelectedItem(ByVal Value As ImcComboItem)
Set Me.SelectedItem = Value
End Property

Public Property Set SelectedItem(ByVal Value As ImcComboItem)
If ImageComboHandle <> 0 Then
    If Not Value Is Nothing Then
        Value.Selected = True
    Else
        SendMessage ImageComboHandle, CB_SETCURSEL, -1, ByVal 0&
    End If
End If
End Property

Public Property Get DroppedDown() As Boolean
Attribute DroppedDown.VB_Description = "Returns/sets a value that determines whether the drop-down list is dropped down or not."
Attribute DroppedDown.VB_MemberFlags = "400"
If ImageComboHandle <> 0 Then DroppedDown = CBool(SendMessage(ImageComboHandle, CB_GETDROPPEDSTATE, 0, ByVal 0&) <> 0)
End Property

Public Property Let DroppedDown(ByVal Value As Boolean)
If ImageComboHandle <> 0 Then SendMessage ImageComboHandle, CB_SHOWDROPDOWN, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Property Get DropDownWidth() As Single
Attribute DropDownWidth.VB_Description = "Returns/sets the width of the drop-down list. This property is not supported in a simple image combo."
Attribute DropDownWidth.VB_MemberFlags = "400"
If ImageComboComboHandle <> 0 Then
    Dim RetVal As Long
    RetVal = SendMessage(ImageComboComboHandle, CB_GETDROPPEDWIDTH, 0, ByVal 0&)
    If Not RetVal = CB_ERR Then
        DropDownWidth = UserControl.ScaleX(RetVal, vbPixels, vbContainerSize)
    Else
        Err.Raise 5
    End If
End If
End Property

Public Property Let DropDownWidth(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If ImageComboHandle <> 0 Then
    If SendMessage(ImageComboHandle, CB_SETDROPPEDWIDTH, CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels)), ByVal 0&) = CB_ERR Then Err.Raise 5
End If
End Property

Public Function FindItem(ByVal Text As String, Optional ByVal Index As Long, Optional ByVal Partial As Boolean, Optional ByVal Wrap As Boolean) As ImcComboItem
Attribute FindItem.VB_Description = "Finds an item in the list and returns a reference to that item."
If ImageComboComboHandle <> 0 Then
    If Index >= 0 Then
        Dim Count As Long
        Count = SendMessage(ImageComboComboHandle, CB_GETCOUNT, 0, ByVal 0&)
        If Count > 0 Then
            If Index <= Count Then
                If Index > 0 Then Index = Index - 1
                Dim Result As Long, Buffer As String, i As Long
                Result = CB_ERR
                For i = Index To (Count - 1)
                    Buffer = String(SendMessage(ImageComboComboHandle, CB_GETLBTEXTLEN, i, ByVal 0&), vbNullChar)
                    SendMessage ImageComboComboHandle, CB_GETLBTEXT, i, ByVal StrPtr(Buffer)
                    If Len(Buffer) > 0 Then
                        If Partial = True Then
                            If InStr(1, Buffer, Text, vbTextCompare) <> 0 Then
                                Result = i
                                Exit For
                            End If
                        Else
                            If StrComp(Buffer, Text, vbTextCompare) = 0 Then
                                Result = i
                                Exit For
                            End If
                        End If
                    End If
                Next i
                If Wrap = True And Result = CB_ERR And Index > 0 Then
                    For i = 0 To (Index - 1)
                        Buffer = String(SendMessage(ImageComboComboHandle, CB_GETLBTEXTLEN, i, ByVal 0&), vbNullChar)
                        SendMessage ImageComboComboHandle, CB_GETLBTEXT, i, ByVal StrPtr(Buffer)
                        If Len(Buffer) > 0 Then
                            If Partial = True Then
                                If InStr(1, Buffer, Text, vbTextCompare) <> 0 Then
                                    Result = i
                                    Exit For
                                End If
                            Else
                                If StrComp(Buffer, Text, vbTextCompare) = 0 Then
                                    Result = i
                                    Exit For
                                End If
                            End If
                        End If
                    Next i
                End If
                If Not Result = CB_ERR Then Set FindItem = Me.ComboItems(Result + 1)
            Else
                Err.Raise 380
            End If
        End If
    Else
        Err.Raise 380
    End If
End If
End Function

Public Property Get OLEDraggedItem() As ImcComboItem
Attribute OLEDraggedItem.VB_Description = "Returns a reference to the currently dragged combo item during an OLE drag/drop operation."
Attribute OLEDraggedItem.VB_MemberFlags = "400"
If ImageComboDragIndex > 0 Then Set OLEDraggedItem = Me.ComboItems(ImageComboDragIndex)
End Property

Private Sub SetDropDownHeight(ByVal Calculate As Boolean)
Static LastCount As Long, ItemHeight As Long, ImageHeight As Long
If ImageComboHandle <> 0 Then
    Dim Count As Long, hImageList As Long
    Count = SendMessage(ImageComboHandle, CB_GETCOUNT, 0, ByVal 0&)
    Select Case Count
        Case 0
            Count = 1
        Case Is > PropMaxDropDownItems
            Count = PropMaxDropDownItems
    End Select
    If Calculate = False Then
        If Count = LastCount Then Exit Sub
    Else
        ItemHeight = SendMessage(ImageComboHandle, CB_GETITEMHEIGHT, 0, ByVal 0&)
        hImageList = SendMessage(ImageComboHandle, CBEM_GETIMAGELIST, 0, ByVal 0&)
        If hImageList <> 0 Then ImageList_GetIconSize hImageList, 0, ImageHeight
    End If
    If ImageComboComboHandle <> 0 Then
        If PropStyle <> ImcStyleSimpleCombo Then
            MoveWindow ImageComboComboHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + (IIf(ImageHeight > ItemHeight, ImageHeight, ItemHeight) * Count) + 2, 1
        Else
            RedrawWindow ImageComboComboHandle, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
        End If
    End If
    LastCount = Count
End If
End Sub

Private Sub CheckTopIndex()
Dim TopIndex As Long
If ImageComboComboHandle <> 0 Then TopIndex = SendMessage(ImageComboComboHandle, CB_GETTOPINDEX, 0, ByVal 0&)
If TopIndex <> ImageComboTopIndex Then
    ImageComboTopIndex = TopIndex
    RaiseEvent Scroll
End If
End Sub

Private Function PropImageListControl() As Object
If ImageComboImageListObjectPointer <> 0 Then Set PropImageListControl = PtrToObj(ImageComboImageListObjectPointer)
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcCombo(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcEdit(hWnd, wMsg, wParam, lParam)
    Case 4
        ISubclass_Message = WindowProcList(hWnd, wMsg, wParam, lParam)
    Case 5
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
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcCombo(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd And wParam <> ImageComboHandle And (wParam <> ImageComboEditHandle Or ImageComboEditHandle = 0) Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_LBUTTONDOWN
        If ImageComboEditHandle = 0 Then
            Select Case GetFocus()
                Case hWnd, ImageComboHandle
                Case Else
                    UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
            End Select
        Else
            Select Case GetFocus()
                Case hWnd, ImageComboHandle, ImageComboEditHandle
                Case Else
                    UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
            End Select
        End If
        PostMessage hWnd, UM_BUTTONDOWN, MakeDWord(vbLeftButton, GetShiftStateFromParam(wParam)), ByVal lParam
    Case WM_RBUTTONDOWN
        PostMessage hWnd, UM_BUTTONDOWN, MakeDWord(vbRightButton, GetShiftStateFromParam(wParam)), ByVal lParam
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        If PropStyle = ImcStyleDropDownList Then
            Dim KeyCode As Integer
            KeyCode = wParam And &HFF&
            If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
                If wMsg = WM_KEYDOWN Then
                    RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
                ElseIf wMsg = WM_KEYUP Then
                    RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
                End If
                ImageComboCharCodeCache = ComCtlsPeekCharCode(hWnd)
            ElseIf wMsg = WM_SYSKEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_SYSKEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            wParam = KeyCode
        End If
    Case WM_CHAR
        If PropStyle = ImcStyleDropDownList Then
            Dim KeyChar As Integer
            If ImageComboCharCodeCache <> 0 Then
                KeyChar = CUIntToInt(ImageComboCharCodeCache And &HFFFF&)
                ImageComboCharCodeCache = 0
            Else
                KeyChar = CUIntToInt(wParam And &HFFFF&)
            End If
            RaiseEvent KeyPress(KeyChar)
            wParam = CIntToUInt(KeyChar)
        End If
    Case WM_UNICHAR
        If PropStyle = ImcStyleDropDownList Then
            If wParam = UNICODE_NOCHAR Then
                WindowProcCombo = 1
            Else
                Dim UTF16 As String
                UTF16 = UTF32CodePoint_To_UTF16(wParam)
                If Len(UTF16) = 1 Then
                    SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
                ElseIf Len(UTF16) = 2 Then
                    SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                    SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
                End If
                WindowProcCombo = 0
            End If
            Exit Function
        End If
    Case WM_IME_CHAR
        If PropStyle = ImcStyleDropDownList Then
            SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
            Exit Function
        End If
    Case WM_CHARTOITEM
        Dim CharCode As Integer, Count As Long, Result As Long
        CharCode = LoWord(wParam)
        Count = SendMessage(hWnd, CB_GETCOUNT, 0, ByVal 0&)
        Result = CB_ERR
        If Count > 0 Then
            Dim CurrIndex As Long, Buffer As String, i As Long
            CurrIndex = SendMessage(hWnd, CB_GETCURSEL, 0, ByVal 0&)
            For i = (CurrIndex + 1) To (Count - 1)
                Buffer = String(SendMessage(hWnd, CB_GETLBTEXTLEN, i, ByVal 0&), vbNullChar)
                SendMessage hWnd, CB_GETLBTEXT, i, ByVal StrPtr(Buffer)
                If Len(Buffer) > 0 Then
                    If StrComp(Left$(Buffer, 1), ChrW(CharCode), vbTextCompare) = 0 Then
                        Result = i
                        Exit For
                    End If
                End If
            Next i
            If Result = CB_ERR And CurrIndex > -1 Then
                For i = 0 To CurrIndex
                    Buffer = String(SendMessage(hWnd, CB_GETLBTEXTLEN, i, ByVal 0&), vbNullChar)
                    SendMessage hWnd, CB_GETLBTEXT, i, ByVal StrPtr(Buffer)
                    If Len(Buffer) > 0 Then
                        If StrComp(Left$(Buffer, 1), ChrW(CharCode), vbTextCompare) = 0 Then
                            Result = i
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If
        WindowProcCombo = Result
        Exit Function
    Case WM_COMMAND
        If PropStyle = ImcStyleDropDownCombo Then
            If lParam = ImageComboEditHandle Then
                Const EN_UPDATE As Long = &H400
                If HiWord(wParam) = EN_UPDATE Then RedrawWindow ImageComboEditHandle, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
            End If
        End If
    Case UM_BUTTONDOWN
        ' The control enters a modal message loop on WM_LBUTTONDOWN and WM_RBUTTONDOWN. (DragDetect)
        ' This workaround is necessary to raise 'MouseDown' before the button was released or the mouse was moved.
        RaiseEvent MouseDown(LoWord(wParam), HiWord(wParam), UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips), UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips))
        Exit Function
End Select
WindowProcCombo = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                ' In case DragDetect returns 0 then the control will set focus the focus automatically.
                ' Otherwise not. So check and change focus, if needed.
                If ImageComboEditHandle = 0 Then
                    Select Case GetFocus()
                        Case hWnd, ImageComboHandle
                        Case Else
                            SetFocusAPI hWnd
                    End Select
                Else
                    Select Case GetFocus()
                        Case hWnd, ImageComboHandle, ImageComboEditHandle
                        Case Else
                            SetFocusAPI ImageComboEditHandle
                    End Select
                End If
                ' See UM_BUTTONDOWN
                If ComCtlsSupportLevel() = 0 Then
                    ' The WM_LBUTTONUP message is not sent if the comctl32.dll version is 5.8x. (bug?)
                    If SendMessage(hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0&) <> 0 Then PostMessage hWnd, WM_LBUTTONUP, wParam, ByVal lParam
                End If
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONDOWN
                ' See UM_BUTTONDOWN
            Case WM_MOUSEMOVE
                If (ImageComboMouseOver(0) = False And PropMouseTrack = True) Or (ImageComboMouseOver(2) = False And PropMouseTrack = True) Then
                    If ImageComboMouseOver(0) = False And PropMouseTrack = True Then ImageComboMouseOver(0) = True
                    If ImageComboMouseOver(2) = False And PropMouseTrack = True Then
                        ImageComboMouseOver(2) = True
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
        ImageComboMouseOver(0) = False
        If ImageComboMouseOver(2) = True Then
            Dim Pos As Long
            Pos = GetMessagePos()
            If WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos)) <> ImageComboEditHandle Or ImageComboEditHandle = 0 Then
                ImageComboMouseOver(2) = False
                RaiseEvent MouseLeave
            End If
        End If
    Case CB_SETTOPINDEX
        Call CheckTopIndex
End Select
End Function

Private Function WindowProcEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        If wParam <> UserControl.hWnd And wParam <> ImageComboHandle And wParam <> ImageComboComboHandle Then SetFocusAPI UserControl.hWnd: Exit Function
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_LBUTTONDOWN
        Select Case GetFocus()
            Case hWnd, ImageComboHandle, ImageComboComboHandle
            Case Else
                UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
        End Select
    Case WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP
        Dim KeyCode As Integer
        KeyCode = wParam And &HFF&
        If wMsg = WM_KEYDOWN Or wMsg = WM_KEYUP Then
            If wMsg = WM_KEYDOWN Then
                RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            ImageComboCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If ImageComboCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(ImageComboCharCodeCache And &HFFFF&)
            ImageComboCharCodeCache = 0
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
    Case WM_INPUTLANGCHANGE
        Call ComCtlsSetIMEMode(hWnd, ImageComboIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call ComCtlsSetIMEMode(hWnd, ImageComboIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case UM_SETFOCUS
        SetFocusAPI hWnd
        Exit Function
End Select
WindowProcEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim P As POINTAPI
        P.X = Get_X_lParam(lParam)
        P.Y = Get_Y_lParam(lParam)
        If ImageComboComboHandle <> 0 Then MapWindowPoints hWnd, ImageComboComboHandle, P, 1
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
                If (ImageComboMouseOver(1) = False And PropMouseTrack = True) Or (ImageComboMouseOver(2) = False And PropMouseTrack = True) Then
                    If ImageComboMouseOver(1) = False And PropMouseTrack = True Then ImageComboMouseOver(1) = True
                    If ImageComboMouseOver(2) = False And PropMouseTrack = True Then
                        ImageComboMouseOver(2) = True
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
        ImageComboMouseOver(1) = False
        If ImageComboMouseOver(2) = True Then
            Dim Pos As Long
            Pos = GetMessagePos()
            If WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos)) <> ImageComboComboHandle Or ImageComboComboHandle = 0 Then
                ImageComboMouseOver(2) = False
                RaiseEvent MouseLeave
            End If
        End If
End Select
End Function

Private Function WindowProcList(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_KEYDOWN, WM_KEYUP
        If PropLocked = True Then
            Dim KeyCode As Integer
            KeyCode = wParam And &HFF&
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
                    Exit Function
            End Select
        End If
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP, WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        If PropLocked = True Then
            Dim P As POINTAPI
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            ClientToScreen hWnd, P
            If Not LBItemFromPt(hWnd, P.X, P.Y, 0) = LB_ERR Then Exit Function
        End If
    Case WM_VSCROLL
        Select Case LoWord(wParam)
            Case SB_THUMBPOSITION, SB_THUMBTRACK
                ' HiWord carries only 16 bits of scroll box position data.
                ' Below workaround will circumvent the 16-bit barrier by using the 32-bit GetScrollInfo function.
                Dim dwStyle As Long
                dwStyle = GetWindowLong(ImageComboListHandle, GWL_STYLE)
                If lParam = 0 And (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                    Dim SCI As SCROLLINFO, PrevPos As Long
                    SCI.cbSize = LenB(SCI)
                    SCI.fMask = SIF_POS Or SIF_TRACKPOS
                    GetScrollInfo ImageComboListHandle, SB_VERT, SCI
                    PrevPos = SCI.nPos
                    Select Case LoWord(wParam)
                        Case SB_THUMBPOSITION
                            SCI.nPos = SCI.nTrackPos
                        Case SB_THUMBTRACK
                            If PropScrollTrack = True Then SCI.nPos = SCI.nTrackPos
                    End Select
                    If PrevPos <> SCI.nPos Then
                        ' SetScrollInfo function not needed as CB_SETTOPINDEX itself will do the scrolling.
                        SendMessage ImageComboComboHandle, CB_SETTOPINDEX, SCI.nPos, ByVal 0&
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

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        Select Case HiWord(wParam)
            Case CBN_SELCHANGE
                SetWindowPos ImageComboHandle, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
                Call CheckTopIndex
                RaiseEvent Click
            Case CBN_DBLCLK
                RaiseEvent DblClick
            Case CBN_EDITCHANGE
                If ImageComboHandle <> 0 Then
                    If LoWord(wParam) = GetWindowLong(ImageComboEditHandle, GWL_ID) Then
                        UserControl.PropertyChanged "Text"
                        On Error Resume Next
                        UserControl.Extender.DataChanged = True
                        On Error GoTo 0
                        RaiseEvent Change
                    End If
                End If
            Case CBN_DROPDOWN
                If PropStyle <> ImcStyleDropDownList And ImageComboEditHandle <> 0 Then
                    If GetCursor() = 0 Then
                        ' The mouse cursor can be hidden when showing the drop-down list upon a change event.
                        ' Reason is that the edit control hides the cursor and a following mouse move will show it again.
                        ' However, the drop-down list will set a mouse capture and thus the cursor keeps hidden.
                        ' Solution is to refresh the cursor by sending a WM_SETCURSOR.
                        Call RefreshMousePointer(lParam)
                    End If
                End If
                RaiseEvent DropDown
                If ImageComboEditHandle <> 0 Then PostMessage ImageComboEditHandle, UM_SETFOCUS, 0, ByVal 0&
            Case CBN_CLOSEUP
                RaiseEvent CloseUp
        End Select
    Case WM_NOTIFY
        Dim Item As ImcComboItem
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = ImageComboHandle Then
            Select Case NM.Code
                Case CBEN_DRAGBEGIN
                    Dim NMCBEDB As NMCBEDRAGBEGIN
                    CopyMemory NMCBEDB, ByVal lParam, LenB(NMCBEDB)
                    If NMCBEDB.iItem = -1 Then
                        Set Item = Me.SelectedItem
                        If Not Item Is Nothing Then
                            ImageComboDragIndexBuffer = Item.Index
                            Dim Button As Integer
                            Button = GetMouseStateFromMsg()
                            If Button = vbLeftButton Then
                                RaiseEvent ItemDrag(Item, vbLeftButton)
                                If PropOLEDragMode = vbOLEDragAutomatic Then Me.OLEDrag
                            ElseIf Button = vbRightButton Then
                                RaiseEvent ItemDrag(Item, vbRightButton)
                            End If
                            ImageComboDragIndexBuffer = 0
                        End If
                    End If
                Case CBEN_BEGINEDIT
                    RaiseEvent BeginEdit
                Case CBEN_ENDEDIT
                    Dim NMCBEEE As NMCBEENDEDIT, NewText As String
                    CopyMemory NMCBEEE, ByVal lParam, LenB(NMCBEEE)
                    NewText = VarToStr(NMCBEEE.szText)
                    NewText = Left$(NewText, InStr(NewText, vbNullChar) - 1)
                    RaiseEvent EndEdit(CBool(NMCBEEE.fChanged <> 0), NMCBEEE.iNewSelection, NewText, NMCBEEE.iWhy)
                Case CBEN_GETDISPINFO
                    Dim NMCBE As NMCOMBOBOXEX
                    CopyMemory NMCBE, ByVal lParam, LenB(NMCBE)
                    With NMCBE.CeItem
                    If .iItem > -1 Then
                        Set Item = Me.ComboItems(.iItem + 1)
                        If (.Mask And CBEIF_IMAGE) = CBEIF_IMAGE Then .iImage = Item.ImageIndex - 1
                        If (.Mask And CBEIF_SELECTEDIMAGE) = CBEIF_SELECTEDIMAGE Then
                            .iSelectedImage = Item.SelImageIndex - 1
                            If .iSelectedImage = I_IMAGECALLBACK Then .iSelectedImage = Item.ImageIndex - 1
                        End If
                    End If
                    End With
                    CopyMemory ByVal lParam, NMCBE, LenB(NMCBE)
            End Select
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI ImageComboHandle
End Function
