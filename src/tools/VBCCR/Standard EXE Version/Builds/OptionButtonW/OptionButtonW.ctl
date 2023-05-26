VERSION 5.00
Begin VB.UserControl OptionButtonW 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "OptionButtonW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "OptionButtonW.ctx":0035
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "OptionButtonW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const ImplementThemedGraphical = True

#If False Then
Private OptImageListAlignmentLeft, OptImageListAlignmentRight, OptImageListAlignmentTop, OptImageListAlignmentBottom, OptImageListAlignmentCenter
Private OptDrawModeNormal, OptDrawModeOwnerDraw
#End If
Private Const BUTTON_IMAGELIST_ALIGN_LEFT As Long = 0
Private Const BUTTON_IMAGELIST_ALIGN_RIGHT As Long = 1
Private Const BUTTON_IMAGELIST_ALIGN_TOP As Long = 2
Private Const BUTTON_IMAGELIST_ALIGN_BOTTOM As Long = 3
Private Const BUTTON_IMAGELIST_ALIGN_CENTER As Long = 4
Public Enum OptImageListAlignmentConstants
OptImageListAlignmentLeft = BUTTON_IMAGELIST_ALIGN_LEFT
OptImageListAlignmentRight = BUTTON_IMAGELIST_ALIGN_RIGHT
OptImageListAlignmentTop = BUTTON_IMAGELIST_ALIGN_TOP
OptImageListAlignmentBottom = BUTTON_IMAGELIST_ALIGN_BOTTOM
OptImageListAlignmentCenter = BUTTON_IMAGELIST_ALIGN_CENTER
End Enum
Public Enum OptDrawModeConstants
OptDrawModeNormal = 0
OptDrawModeOwnerDraw = 1
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
Private Type BUTTON_IMAGELIST
hImageList As Long
RCMargin As RECT
uAlign As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMBCHOTITEM
hdr As NMHDR
dwFlags As Long
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
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event HotChanged()
Attribute HotChanged.VB_Description = "Occurrs when the option button control's hot state changes. Requires comctl32.dll version 6.0 or higher."
Public Event OwnerDraw(ByVal Action As Long, ByVal State As Long, ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Attribute OwnerDraw.VB_Description = "Occurs when a visual aspect of an owner-drawn button has changed."
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
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lData As Long, ByVal wData As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal fFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal nCtlType As Long, ByVal nFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As Long) As Long

#If ImplementThemedGraphical = True Then

Private Enum UxThemeButtonParts
BP_PUSHBUTTON = 1
BP_RADIOBUTTON = 2
BP_CHECKBOX = 3
BP_GROUPBOX = 4
BP_USERBUTTON = 5
End Enum
Private Enum UxThemeButtonStates
PBS_NORMAL = 1
PBS_HOT = 2
PBS_PRESSED = 3
PBS_DISABLED = 4
PBS_DEFAULTED = 5
End Enum
Private Const DTT_TEXTCOLOR As Long = 1
Private Type DTTOPTS
dwSize As Long
dwFlags As Long
crText As Long
crBorder As Long
crShadow As Long
eTextShadowType As Long
PTShadowOffset As POINTAPI
iBorderSize As Long
iFontPropId As Long
iColorPropId As Long
iStateId As Long
fApplyOverlay As Long
iGlowSize As Long
End Type
Private Declare Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme" (ByVal Theme As Long, iPartId As Long, iStateId As Long) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme" (ByVal hWnd As Long, ByVal hDC As Long, ByRef pRect As RECT) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef pClipRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, ByRef pRect As RECT) As Long
Private Declare Function DrawThemeTextEx Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByRef lpRect As RECT, ByRef lpOptions As DTTOPTS) As Long
Private Declare Function GetThemeBackgroundRegion Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef hRgn As Long) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pBoundingRect As RECT, ByRef pContentRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal Theme As Long) As Long

#End If

Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const HWND_DESKTOP As Long = &H0
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
Private Const WM_NOTIFY As Long = &H4E
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
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_COMMAND As Long = &H111
Private Const WM_DRAWITEM As Long = &H2B, ODT_BUTTON As Long = &H4, ODA_FOCUS As Long = &H4, ODS_SELECTED As Long = &H1, ODS_DISABLED As Long = &H4, ODS_FOCUS As Long = &H10, ODS_NOACCEL As Long = &H100, ODS_NOFOCUSRECT As Long = &H200
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_CTLCOLORBTN As Long = &H135
Private Const WM_PAINT As Long = &HF
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const DFC_BUTTON As Long = &H4, DFCS_BUTTONPUSH As Long = &H10, DFCS_INACTIVE As Long = &H100, DFCS_PUSHED As Long = &H200, DFCS_CHECKED As Long = &H400, DFCS_ADJUSTRECT As Long = &H2000, DFCS_FLAT As Long = &H4000
Private Const BS_TEXT As Long = &H0
Private Const BS_OWNERDRAW As Long = &HB
Private Const BS_RADIOBUTTON As Long = &H4
Private Const BS_RIGHTBUTTON As Long = &H20
Private Const BS_ICON As Long = &H40
Private Const BS_BITMAP As Long = &H80
Private Const BS_LEFT As Long = &H100
Private Const BS_RIGHT As Long = &H200
Private Const BS_CENTER As Long = &H300
Private Const BS_TOP As Long = &H400
Private Const BS_VCENTER As Long = &HC00
Private Const BS_BOTTOM As Long = &H800
Private Const BS_PUSHLIKE As Long = &H1000
Private Const BS_MULTILINE As Long = &H2000
Private Const BS_NOTIFY As Long = &H4000
Private Const BS_FLAT As Long = &H8000&
Private Const BM_GETCHECK As Long = &HF0
Private Const BM_SETCHECK As Long = &HF1
Private Const BM_GETSTATE As Long = &HF2
Private Const BM_SETSTATE As Long = &HF3
Private Const BM_GETIMAGE As Long = &HF6
Private Const BM_SETIMAGE As Long = &HF7
Private Const WM_USER As Long = &H400
Private Const UM_CHECKVALUE As Long = (WM_USER + 300)
Private Const BCM_FIRST As Long = &H1600
Private Const BCM_SETIMAGELIST As Long = (BCM_FIRST + 2)
Private Const BCM_GETIMAGELIST As Long = (BCM_FIRST + 3)
Private Const BST_UNCHECKED As Long = &H0
Private Const BST_CHECKED As Long = &H1
Private Const BST_PUSHED As Long = &H4
Private Const BST_HOT As Long = &H200
Private Const BCCL_NOGLYPH As Long = (-1) ' Contrary to MSDN it even works on Windows XP
Private Const BN_CLICKED As Long = 0
Private Const BN_DOUBLECLICKED As Long = 5
Private Const BCN_FIRST As Long = -1250
Private Const BCN_HOTITEMCHANGE As Long = (BCN_FIRST + 1)
Private Const HICF_MOUSE As Long = &H1
Private Const HICF_ENTERING As Long = &H10
Private Const HICF_LEAVING As Long = &H20
Private Const IMAGE_BITMAP As Long = 0
Private Const IMAGE_ICON As Long = 1
Private Const DT_CENTER As Long = &H1
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_CALCRECT As Long = &H400
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const RGN_DIFF As Long = 4
Private Const RGN_COPY As Long = 5
Private Const DST_ICON As Long = &H3
Private Const DST_BITMAP As Long = &H4
Private Const DSS_DISABLED As Long = &H20
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private OptionButtonHandle As Long
Private OptionButtonTransparentBrush As Long
Private OptionButtonOwnerDrawCheckedBrush As Long
Private OptionButtonFontHandle As Long
Private OptionButtonCharCodeCache As Long
Private OptionButtonMouseOver(0 To 1) As Boolean
Private OptionButtonDesignMode As Boolean
Private OptionButtonImageListHandle As Long
Private OptionButtonImageListObjectPointer As Long
Private OptionButtonEnabledVisualStyles As Boolean
Private OptionButtonPictureRenderFlag As Integer
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private DispIDImageList As Long, ImageListArray() As String
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropImageListName As String, PropImageListInit As Boolean
Private PropImageListAlignment As OptImageListAlignmentConstants
Private PropImageListMargin As Long
Private PropValue As OLE_OPTEXCLUSIVE
Private PropCaption As String
Private PropAlignment As CCLeftRightAlignmentConstants
Private PropTextAlignment As VBRUN.AlignmentConstants
Private PropPushLike As Boolean
Private PropPicture As IPictureDisp
Private PropWordWrap As Boolean
Private PropTransparent As Boolean
Private PropVerticalAlignment As CCVerticalAlignmentConstants
Private PropStyle As VBRUN.ButtonConstants
Private PropDisabledPicture As IPictureDisp
Private PropDownPicture As IPictureDisp
Private PropUseMaskColor As Boolean
Private PropMaskColor As OLE_COLOR
Private PropDrawMode As OptDrawModeConstants

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
Call ComCtlsInitCC(ICC_STANDARD_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim ImageListArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
OptionButtonDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = "(None)"
If PropRightToLeft = False Then PropImageListAlignment = OptImageListAlignmentLeft Else PropImageListAlignment = OptImageListAlignmentRight
PropImageListMargin = 0
PropValue = False
PropCaption = Ambient.DisplayName
If PropRightToLeft = False Then PropAlignment = CCLeftRightAlignmentLeft Else PropAlignment = CCLeftRightAlignmentRight
If PropRightToLeft = False Then PropTextAlignment = vbLeftJustify Else PropTextAlignment = vbRightJustify
PropPushLike = False
Set PropPicture = Nothing
PropWordWrap = True
PropTransparent = False
PropVerticalAlignment = CCVerticalAlignmentCenter
PropStyle = vbButtonStandard
Set PropDisabledPicture = Nothing
Set PropDownPicture = Nothing
PropUseMaskColor = False
PropMaskColor = &HC0C0C0
PropDrawMode = OptDrawModeNormal
Call CreateOptionButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
OptionButtonDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Appearance = .ReadProperty("Appearance", CCAppearance3D)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
Me.ForeColor = .ReadProperty("ForeColor", vbButtonText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = .ReadProperty("ImageList", "(None)")
PropImageListAlignment = .ReadProperty("ImageListAlignment", OptImageListAlignmentLeft)
PropImageListMargin = .ReadProperty("ImageListMargin", 0)
PropValue = .ReadProperty("Value", False)
PropCaption = .ReadProperty("Caption", vbNullString) ' Unicode not necessary
PropAlignment = .ReadProperty("Alignment", CCLeftRightAlignmentLeft)
PropTextAlignment = .ReadProperty("TextAlignment", vbLeftJustify)
PropPushLike = .ReadProperty("PushLike", False)
Set PropPicture = .ReadProperty("Picture", Nothing)
PropWordWrap = .ReadProperty("WordWrap", True)
PropTransparent = .ReadProperty("Transparent", False)
PropVerticalAlignment = .ReadProperty("VerticalAlignment", CCVerticalAlignmentCenter)
PropStyle = .ReadProperty("Style", vbButtonStandard)
Set PropDisabledPicture = .ReadProperty("DisabledPicture", Nothing)
Set PropDownPicture = .ReadProperty("DownPicture", Nothing)
PropUseMaskColor = .ReadProperty("UseMaskColor", False)
PropMaskColor = .ReadProperty("MaskColor", &HC0C0C0)
PropDrawMode = .ReadProperty("DrawMode", OptDrawModeNormal)
End With
Call CreateOptionButton
If PropValue = True And OptionButtonDesignMode = False Then PostMessage UserControl.hWnd, UM_CHECKVALUE, 0, ByVal 0&
If Not PropImageListName = "(None)" Then TimerImageList.Enabled = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Appearance", Me.Appearance, CCAppearance3D
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "ForeColor", Me.ForeColor, vbButtonText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "ImageList", PropImageListName, "(None)"
.WriteProperty "ImageListAlignment", PropImageListAlignment, OptImageListAlignmentLeft
.WriteProperty "ImageListMargin", PropImageListMargin, 0
.WriteProperty "Value", PropValue, False
.WriteProperty "Caption", PropCaption, vbNullString ' Unicode not necessary
.WriteProperty "Alignment", PropAlignment, CCLeftRightAlignmentLeft
.WriteProperty "TextAlignment", PropTextAlignment, vbLeftJustify
.WriteProperty "PushLike", PropPushLike, False
.WriteProperty "Picture", PropPicture, Nothing
.WriteProperty "WordWrap", PropWordWrap, True
.WriteProperty "Transparent", PropTransparent, False
.WriteProperty "VerticalAlignment", PropVerticalAlignment, CCVerticalAlignmentCenter
.WriteProperty "Style", PropStyle, vbButtonStandard
.WriteProperty "DisabledPicture", PropDisabledPicture, Nothing
.WriteProperty "DownPicture", PropDownPicture, Nothing
.WriteProperty "UseMaskColor", PropUseMaskColor, False
.WriteProperty "MaskColor", PropMaskColor, &HC0C0C0
.WriteProperty "DrawMode", PropDrawMode, OptDrawModeNormal
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
If OptionButtonHandle <> 0 Then
    If PropTransparent = True Then
        MoveWindow OptionButtonHandle, 0, 0, .ScaleWidth, .ScaleHeight, 0
        If OptionButtonTransparentBrush <> 0 Then
            DeleteObject OptionButtonTransparentBrush
            OptionButtonTransparentBrush = 0
        End If
        RedrawWindow OptionButtonHandle, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
    Else
        MoveWindow OptionButtonHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    End If
End If
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyOptionButton
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropImageListInit = False Then
    Me.ImageList = PropImageListName
    PropImageListInit = True
End If
TimerImageList.Enabled = False
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
hWnd = OptionButtonHandle
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
OldFontHandle = OptionButtonFontHandle
OptionButtonFontHandle = CreateGDIFontFromOLEFont(PropFont)
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, WM_SETFONT, OptionButtonFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = OptionButtonFontHandle
OptionButtonFontHandle = CreateGDIFontFromOLEFont(PropFont)
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, WM_SETFONT, OptionButtonFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
OptionButtonEnabledVisualStyles = EnabledVisualStyles()
If OptionButtonHandle <> 0 And OptionButtonEnabledVisualStyles = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles OptionButtonHandle
    Else
        RemoveVisualStyles OptionButtonHandle
    End If
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Appearance() As CCAppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets a value that determines whether an object is painted two-dimensional or with 3-D effects."
Attribute Appearance.VB_UserMemId = -520
Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal Value As CCAppearanceConstants)
Select Case Value
    Case CCAppearanceFlat, CCAppearance3D
        UserControl.Appearance = Value
    Case Else
        Err.Raise 380
End Select
UserControl.ForeColor = IIf(UserControl.Appearance = CCAppearanceFlat, vbWindowText, vbButtonText)
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
        If UserControl.Appearance = CCAppearanceFlat Then
            If Not (dwStyle And BS_FLAT) = BS_FLAT Then dwStyle = dwStyle Or BS_FLAT
        Else
            If (dwStyle And BS_FLAT) = BS_FLAT Then dwStyle = dwStyle And Not BS_FLAT
        End If
        SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
    End If
End If
Me.Refresh
UserControl.PropertyChanged "Appearance"
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
If OptionButtonHandle <> 0 Then EnableWindow OptionButtonHandle, IIf(Value = True, 1, 0)
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
If OptionButtonDesignMode = False Then Call RefreshMousePointer
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
        If OptionButtonDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If OptionButtonDesignMode = False Then Call RefreshMousePointer
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
If OptionButtonHandle <> 0 Then
    Call ComCtlsSetRightToLeft(OptionButtonHandle, dwMask)
    If PropRightToLeft = False Then
        If PropImageListAlignment = OptImageListAlignmentRight Then Me.ImageListAlignment = OptImageListAlignmentLeft
        If PropTextAlignment = vbRightJustify Then Me.TextAlignment = vbLeftJustify
    Else
        If PropImageListAlignment = OptImageListAlignmentLeft Then Me.ImageListAlignment = OptImageListAlignmentRight
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

Public Property Get ImageList() As Variant
Attribute ImageList.VB_Description = "Returns/sets the image list control to be used. The image list should contain either a single image to be used for all states or individual images for each state. Requires comctl32.dll version 6.0 or higher."
If OptionButtonDesignMode = False Then
    If PropImageListInit = False And OptionButtonImageListObjectPointer = 0 Then
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
If OptionButtonHandle <> 0 Then
    ' The image list should contain either a single image to be used for all states or
    ' individual images for each state. The following states are defined as following:
    ' PBS_NORMAL = 1
    ' PBS_HOT = 2
    ' PBS_PRESSED = 3
    ' PBS_DISABLED = 4
    ' PBS_DEFAULTED = 5
    ' PBS_STYLUSHOT = 6
    Dim Success As Boolean, Handle As Long
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> 0)
        End If
        If Success = True Then
            Call SetImageList(Handle)
            OptionButtonImageListObjectPointer = ObjPtr(Value)
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
                        Call SetImageList(Handle)
                        If OptionButtonDesignMode = False Then OptionButtonImageListObjectPointer = ObjPtr(ControlEnum)
                        PropImageListName = Value
                        Exit For
                    ElseIf OptionButtonDesignMode = True Then
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
        Call SetImageList(BCCL_NOGLYPH)
        OptionButtonImageListObjectPointer = 0
        PropImageListName = "(None)"
    ElseIf Handle = 0 Then
        Call SetImageList(BCCL_NOGLYPH)
    End If
End If
UserControl.PropertyChanged "ImageList"
End Property

Public Property Get ImageListAlignment() As OptImageListAlignmentConstants
Attribute ImageListAlignment.VB_Description = "Returns/sets the alignment used to the image in the image list control. Requires comctl32.dll version 6.0 or higher."
ImageListAlignment = PropImageListAlignment
End Property

Public Property Let ImageListAlignment(ByVal Value As OptImageListAlignmentConstants)
Select Case Value
    Case OptImageListAlignmentLeft, OptImageListAlignmentRight, OptImageListAlignmentTop, OptImageListAlignmentBottom, OptImageListAlignmentCenter
        PropImageListAlignment = Value
    Case Else
        Err.Raise 380
End Select
If OptionButtonHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If Not PropImageListControl Is Nothing Then
        Me.ImageList = PropImageListControl
    ElseIf Not PropImageListName = "(None)" Then
        Me.ImageList = PropImageListName
    End If
End If
UserControl.PropertyChanged "ImageListAlignment"
End Property

Public Property Get ImageListMargin() As Single
Attribute ImageListMargin.VB_Description = "Returns/sets the margin (related to the alignment) used to the image in the image list control. Requires comctl32.dll version 6.0 or higher."
ImageListMargin = UserControl.ScaleX(PropImageListMargin, vbPixels, vbContainerSize)
End Property

Public Property Let ImageListMargin(ByVal Value As Single)
If Value < 0 Then
    If OptionButtonDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropImageListMargin = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If OptionButtonHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If Not PropImageListControl Is Nothing Then
        Me.ImageList = PropImageListControl
    ElseIf Not PropImageListName = "(None)" Then
        Me.ImageList = PropImageListName
    End If
End If
UserControl.PropertyChanged "ImageListMargin"
End Property

Public Property Get Value() As OLE_OPTEXCLUSIVE
Attribute Value.VB_Description = "Returns/sets the value of an object."
Attribute Value.VB_UserMemId = 0
Value = PropValue
End Property

Public Property Let Value(ByVal NewValue As OLE_OPTEXCLUSIVE)
Dim Changed As Boolean
Changed = CBool(PropValue <> NewValue)
PropValue = NewValue
If OptionButtonHandle <> 0 Then
    If Not (GetWindowLong(OptionButtonHandle, GWL_STYLE) And BS_OWNERDRAW) = BS_OWNERDRAW Then
        SendMessage OptionButtonHandle, BM_SETCHECK, IIf(PropValue = True, BST_CHECKED, BST_UNCHECKED), ByVal 0&
    Else
        Me.Refresh
    End If
End If
UserControl.PropertyChanged "Value"
If Changed = True And PropValue = True Then RaiseEvent Click
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
If OptionButtonHandle <> 0 Then
    Caption = String(SendMessage(OptionButtonHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage OptionButtonHandle, WM_GETTEXT, Len(Caption) + 1, ByVal StrPtr(Caption)
Else
    Caption = PropCaption
End If
End Property

Public Property Let Caption(ByVal Value As String)
PropCaption = Value
UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, WM_SETTEXT, 0, ByVal StrPtr(PropCaption)
UserControl.PropertyChanged "Caption"
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
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
        If PropAlignment = CCLeftRightAlignmentRight Then
            If Not (dwStyle And BS_RIGHTBUTTON) = BS_RIGHTBUTTON Then dwStyle = dwStyle Or BS_RIGHTBUTTON
        ElseIf PropAlignment = CCLeftRightAlignmentLeft Then
            If (dwStyle And BS_RIGHTBUTTON) = BS_RIGHTBUTTON Then dwStyle = dwStyle And Not BS_RIGHTBUTTON
        End If
        SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
        Me.Refresh
    End If
End If
UserControl.PropertyChanged "Alignment"
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
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
        If (dwStyle And BS_LEFT) = BS_LEFT Then dwStyle = dwStyle And Not BS_LEFT
        If (dwStyle And BS_CENTER) = BS_CENTER Then dwStyle = dwStyle And Not BS_CENTER
        If (dwStyle And BS_RIGHT) = BS_RIGHT Then dwStyle = dwStyle And Not BS_RIGHT
        Select Case PropTextAlignment
            Case vbLeftJustify
                dwStyle = dwStyle Or BS_LEFT
            Case vbCenter
                dwStyle = dwStyle Or BS_CENTER
            Case vbRightJustify
                dwStyle = dwStyle Or BS_RIGHT
        End Select
        SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
        Me.Refresh
    End If
End If
UserControl.PropertyChanged "TextAlignment"
End Property

Public Property Get PushLike() As Boolean
Attribute PushLike.VB_Description = "Returns/sets a value that determines whether or not the control look and act like a push button."
PushLike = PropPushLike
End Property

Public Property Let PushLike(ByVal Value As Boolean)
PropPushLike = Value
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
        If PropPushLike = True Then
            If Not (dwStyle And BS_PUSHLIKE) = BS_PUSHLIKE Then dwStyle = dwStyle Or BS_PUSHLIKE
        Else
            If (dwStyle And BS_PUSHLIKE) = BS_PUSHLIKE Then dwStyle = dwStyle And Not BS_PUSHLIKE
        End If
        SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
        Me.Refresh
    End If
End If
UserControl.PropertyChanged "PushLike"
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Set Me.Picture = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
Dim dwStyle As Long
If Value Is Nothing Then
    Set PropPicture = Nothing
    If OptionButtonHandle <> 0 And OptionButtonImageListHandle = 0 Then
        dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
        If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
            If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
            If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
            SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
            SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
            SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
            Me.Refresh
        End If
    End If
Else
    Set UserControl.Picture = Value
    Set PropPicture = UserControl.Picture
    Set UserControl.Picture = Nothing
    If OptionButtonHandle <> 0 And OptionButtonImageListHandle = 0 Then
        dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
        If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
            If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
            If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
            If PropPicture.Handle <> 0 Then
                If PropPicture.Type = vbPicTypeIcon Then
                    dwStyle = dwStyle Or BS_ICON
                    SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
                    SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
                    SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal PropPicture.Handle
                Else
                    dwStyle = dwStyle Or BS_BITMAP
                    SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
                    SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
                    SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal PropPicture.Handle
                End If
            Else
                SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
                SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
                SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
            End If
            Me.Refresh
        End If
    End If
End If
If dwStyle = 0 Then dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
OptionButtonPictureRenderFlag = 0
If (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then Me.Refresh
UserControl.PropertyChanged "Picture"
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control may break lines within the text in order to prevent overflow."
WordWrap = PropWordWrap
End Property

Public Property Let WordWrap(ByVal Value As Boolean)
PropWordWrap = Value
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
        If PropWordWrap = True Then
            If Not (dwStyle And BS_MULTILINE) = BS_MULTILINE Then dwStyle = dwStyle Or BS_MULTILINE
        Else
            If (dwStyle And BS_MULTILINE) = BS_MULTILINE Then dwStyle = dwStyle And Not BS_MULTILINE
        End If
        SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
        Me.Refresh
    End If
End If
UserControl.PropertyChanged "WordWrap"
End Property

Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Returns/sets a value indicating if the background is a replica of the underlying background to simulate transparency. This property is ignored at design time."
Transparent = PropTransparent
End Property

Public Property Let Transparent(ByVal Value As Boolean)
PropTransparent = Value
Me.Refresh
UserControl.PropertyChanged "Transparent"
End Property

Public Property Get VerticalAlignment() As CCVerticalAlignmentConstants
Attribute VerticalAlignment.VB_Description = "Returns/sets the vertical alignment."
VerticalAlignment = PropVerticalAlignment
End Property

Public Property Let VerticalAlignment(ByVal Value As CCVerticalAlignmentConstants)
Select Case Value
    Case CCVerticalAlignmentTop, CCVerticalAlignmentCenter, CCVerticalAlignmentBottom
        PropVerticalAlignment = Value
    Case Else
        Err.Raise 380
End Select
If OptionButtonHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
    If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
        If (dwStyle And BS_TOP) = BS_TOP Then dwStyle = dwStyle And Not BS_TOP
        If (dwStyle And BS_VCENTER) = BS_VCENTER Then dwStyle = dwStyle And Not BS_VCENTER
        If (dwStyle And BS_BOTTOM) = BS_BOTTOM Then dwStyle = dwStyle And Not BS_BOTTOM
        Select Case PropVerticalAlignment
            Case CCVerticalAlignmentTop
                dwStyle = dwStyle Or BS_TOP
            Case CCVerticalAlignmentCenter
                dwStyle = dwStyle Or BS_VCENTER
            Case CCVerticalAlignmentBottom
                dwStyle = dwStyle Or BS_BOTTOM
        End Select
        SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
        Me.Refresh
    End If
End If
UserControl.PropertyChanged "VerticalAlignment"
End Property

Public Property Get Style() As VBRUN.ButtonConstants
Attribute Style.VB_Description = "Returns/sets the appearance of the control, whether standard or graphical."
Style = PropStyle
End Property

Public Property Let Style(ByVal Value As VBRUN.ButtonConstants)
Select Case Value
    Case vbButtonStandard, vbButtonGraphical
        If PropDrawMode <> OptDrawModeNormal And Value = vbButtonGraphical Then
            If OptionButtonDesignMode = True Then
                MsgBox "Style must be 0 - Standard when DrawMode is not 0 - Normal", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise Number:=383, Description:="Style must be 0 - Standard when DrawMode is not 0 - Normal"
            End If
        End If
        PropStyle = Value
    Case Else
        Err.Raise 380
End Select
If OptionButtonHandle <> 0 Then Call ReCreateOptionButton
UserControl.PropertyChanged "Style"
End Property

Public Property Get DisabledPicture() As IPictureDisp
Attribute DisabledPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is disabled. Only applicable if the style property is set to 1."
Set DisabledPicture = PropDisabledPicture
End Property

Public Property Let DisabledPicture(ByVal Value As IPictureDisp)
Set Me.DisabledPicture = Value
End Property

Public Property Set DisabledPicture(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropDisabledPicture = Nothing
Else
    Set UserControl.Picture = Value
    Set PropDisabledPicture = UserControl.Picture
    Set UserControl.Picture = Nothing
End If
OptionButtonPictureRenderFlag = 0
Me.Refresh
UserControl.PropertyChanged "DisabledPicture"
End Property

Public Property Get DownPicture() As IPictureDisp
Attribute DownPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is in the down position. Only applicable if the style property is set to 1."
Set DownPicture = PropDownPicture
End Property

Public Property Let DownPicture(ByVal Value As IPictureDisp)
Set Me.DownPicture = Value
End Property

Public Property Set DownPicture(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropDownPicture = Nothing
Else
    Set UserControl.Picture = Value
    Set PropDownPicture = UserControl.Picture
    Set UserControl.Picture = Nothing
End If
OptionButtonPictureRenderFlag = 0
Me.Refresh
UserControl.PropertyChanged "DownPicture"
End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value which determines if the button control will use the mask color property. Only applicable if the style property is set to 1."
UseMaskColor = PropUseMaskColor
End Property

Public Property Let UseMaskColor(ByVal Value As Boolean)
PropUseMaskColor = Value
Me.Refresh
UserControl.PropertyChanged "UseMaskColor"
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in a picture to be a 'mask' (that is, transparent). Only applicable if the style property is set to 1."
MaskColor = PropMaskColor
End Property

Public Property Let MaskColor(ByVal Value As OLE_COLOR)
PropMaskColor = Value
Me.Refresh
UserControl.PropertyChanged "MaskColor"
End Property

Public Property Get DrawMode() As OptDrawModeConstants
Attribute DrawMode.VB_Description = "Returns/sets a value indicating whether your code or the operating system will handle drawing of the elements."
DrawMode = PropDrawMode
End Property

Public Property Let DrawMode(ByVal Value As OptDrawModeConstants)
Select Case Value
    Case OptDrawModeNormal, OptDrawModeOwnerDraw
        PropDrawMode = Value
    Case Else
        Err.Raise 380
End Select
If OptionButtonHandle <> 0 Then Call ReCreateOptionButton
UserControl.PropertyChanged "DrawMode"
End Property

Private Sub CreateOptionButton()
If OptionButtonHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or BS_RADIOBUTTON Or BS_TEXT Or BS_NOTIFY
If Me.Appearance = CCAppearanceFlat Then dwStyle = dwStyle Or BS_FLAT
If PropRightToLeft = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
If PropAlignment = CCLeftRightAlignmentRight Then dwStyle = dwStyle Or BS_RIGHTBUTTON
Select Case PropTextAlignment
    Case vbLeftJustify
        dwStyle = dwStyle Or BS_LEFT
    Case vbCenter
        dwStyle = dwStyle Or BS_CENTER
    Case vbRightJustify
        dwStyle = dwStyle Or BS_RIGHT
End Select
If PropPushLike = True Then dwStyle = dwStyle Or BS_PUSHLIKE
If PropWordWrap = True Then dwStyle = dwStyle Or BS_MULTILINE
Select Case PropVerticalAlignment
    Case CCVerticalAlignmentTop
        dwStyle = dwStyle Or BS_TOP
    Case CCVerticalAlignmentCenter
        dwStyle = dwStyle Or BS_VCENTER
    Case CCVerticalAlignmentBottom
        dwStyle = dwStyle Or BS_BOTTOM
End Select
If PropDrawMode <> OptDrawModeNormal Then PropStyle = vbButtonStandard
If PropStyle = vbButtonGraphical Then dwStyle = dwStyle Or BS_OWNERDRAW
If PropDrawMode = OptDrawModeOwnerDraw Then dwStyle = dwStyle Or BS_OWNERDRAW
If (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
    ' According to MSDN:
    ' The BS_OWNERDRAW style cannot be combined with any other button style.
    dwStyle = WS_CHILD Or WS_VISIBLE Or BS_OWNERDRAW
End If
OptionButtonHandle = CreateWindowEx(dwExStyle, StrPtr("Button"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If OptionButtonHandle <> 0 Then
    Call ComCtlsShowAllUIStates(OptionButtonHandle)
    If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then SendMessage OptionButtonHandle, BM_SETCHECK, IIf(PropValue = True, BST_CHECKED, BST_UNCHECKED), ByVal 0&
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Caption = PropCaption
If Not PropPicture Is Nothing Then Set Me.Picture = PropPicture
If OptionButtonDesignMode = False Then
    If OptionButtonHandle <> 0 Then Call ComCtlsSetSubclass(OptionButtonHandle, Me, 1)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
Else
    If PropStyle = vbButtonGraphical Then
        Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
        Me.Refresh
    End If
End If
End Sub

Private Sub ReCreateOptionButton()
If OptionButtonDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Call DestroyOptionButton
    Call CreateOptionButton
    Call UserControl_Resize
    If Not PropImageListControl Is Nothing Then Set Me.ImageList = PropImageListControl
    If Locked = True Then LockWindowUpdate 0
    Me.Refresh
Else
    Call DestroyOptionButton
    Call CreateOptionButton
    Call UserControl_Resize
    If Not PropImageListName = "(None)" Then Me.ImageList = PropImageListName
End If
End Sub

Private Sub DestroyOptionButton()
If OptionButtonHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(OptionButtonHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow OptionButtonHandle, SW_HIDE
SetParent OptionButtonHandle, 0
DestroyWindow OptionButtonHandle
OptionButtonHandle = 0
If OptionButtonFontHandle <> 0 Then
    DeleteObject OptionButtonFontHandle
    OptionButtonFontHandle = 0
End If
If OptionButtonTransparentBrush <> 0 Then
    DeleteObject OptionButtonTransparentBrush
    OptionButtonTransparentBrush = 0
End If
If OptionButtonOwnerDrawCheckedBrush <> 0 Then
    DeleteObject OptionButtonOwnerDrawCheckedBrush
    OptionButtonOwnerDrawCheckedBrush = 0
End If
OptionButtonImageListHandle = 0
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
If OptionButtonTransparentBrush <> 0 Then
    DeleteObject OptionButtonTransparentBrush
    OptionButtonTransparentBrush = 0
End If
If OptionButtonOwnerDrawCheckedBrush <> 0 Then
    DeleteObject OptionButtonOwnerDrawCheckedBrush
    OptionButtonOwnerDrawCheckedBrush = 0
End If
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get Pushed() As Boolean
Attribute Pushed.VB_Description = "Returns/sets a value that indicates if the option button is in the pushed state."
Attribute Pushed.VB_MemberFlags = "400"
If OptionButtonHandle <> 0 Then Pushed = CBool((SendMessage(OptionButtonHandle, BM_GETSTATE, 0, ByVal 0&) And BST_PUSHED) = BST_PUSHED)
End Property

Public Property Let Pushed(ByVal Value As Boolean)
If OptionButtonHandle <> 0 Then SendMessage OptionButtonHandle, BM_SETSTATE, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Returns/sets a value that indicates if the option button is hot; that is, the mouse is hovering over it. Requires comctl32.dll version 6.0 or higher."
Attribute Hot.VB_MemberFlags = "400"
If OptionButtonHandle <> 0 And ComCtlsSupportLevel() >= 1 Then Hot = CBool((SendMessage(OptionButtonHandle, BM_GETSTATE, 0, ByVal 0&) And BST_HOT) = BST_HOT)
End Property

Public Property Let Hot(ByVal Value As Boolean)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Private Sub SetImageList(ByVal hImageList As Long)
If OptionButtonHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim BTNIML As BUTTON_IMAGELIST
    With BTNIML
    .hImageList = hImageList
    If .hImageList = 0 Then .hImageList = BCCL_NOGLYPH
    OptionButtonImageListHandle = hImageList
    If OptionButtonImageListHandle = BCCL_NOGLYPH Then OptionButtonImageListHandle = 0
    If .hImageList <> BCCL_NOGLYPH Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(OptionButtonHandle, GWL_STYLE)
        If Not (dwStyle And BS_OWNERDRAW) = BS_OWNERDRAW Then
            If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
            If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
            SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
            SendMessage OptionButtonHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
            SetWindowLong OptionButtonHandle, GWL_STYLE, dwStyle
        End If
    End If
    With .RCMargin
    Select Case PropImageListAlignment
        Case OptImageListAlignmentLeft
            .Left = PropImageListMargin
        Case OptImageListAlignmentRight
            .Right = PropImageListMargin
        Case OptImageListAlignmentTop
            .Top = PropImageListMargin
        Case OptImageListAlignmentBottom
            .Bottom = PropImageListMargin
    End Select
    End With
    .uAlign = PropImageListAlignment
    SendMessage OptionButtonHandle, BCM_SETIMAGELIST, 0, ByVal VarPtr(BTNIML)
    If .hImageList = BCCL_NOGLYPH Then Set Me.Picture = PropPicture
    End With
    Me.Refresh
End If
End Sub

Private Sub OffsetRect(ByRef RC As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
With RC
.Left = .Left + X1
.Top = .Top + Y1
.Right = .Right + X2
.Bottom = .Bottom + Y2
End With
End Sub

Private Function CoalescePicture(ByVal Picture As IPictureDisp, ByVal DefaultPicture As IPictureDisp) As IPictureDisp
If Picture Is Nothing Then
    Set CoalescePicture = DefaultPicture
ElseIf Picture.Handle = 0 Then
    Set CoalescePicture = DefaultPicture
Else
    Set CoalescePicture = Picture
End If
End Function

Private Function PropImageListControl() As Object
If OptionButtonImageListObjectPointer <> 0 Then Set PropImageListControl = PtrToObj(OptionButtonImageListObjectPointer)
End Function

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
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            OptionButtonCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If OptionButtonCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(OptionButtonCharCodeCache And &HFFFF&)
            OptionButtonCharCodeCache = 0
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
    Case WM_LBUTTONDBLCLK
        If (GetWindowLong(hWnd, GWL_STYLE) And BS_OWNERDRAW) = BS_OWNERDRAW Then
            ' Buttons having the BS_OWNERDRAW style will not respond to double click as normal buttons do.
            ' Thus the default window procedure of the button will be called with WM_LBUTTONDOWN instead of the actual WM_LBUTTONDBLCLK.
            WindowProcControl = ComCtlsDefaultProc(hWnd, WM_LBUTTONDOWN, wParam, lParam)
            Exit Function
        End If
    
    #If ImplementThemedGraphical = True Then
    
    Case WM_THEMECHANGED
        OptionButtonEnabledVisualStyles = EnabledVisualStyles()
    
    #End If
    
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
                If (OptionButtonMouseOver(0) = False And PropStyle = vbButtonGraphical) Or (OptionButtonMouseOver(1) = False And PropMouseTrack = True) Then
                    
                    #If ImplementThemedGraphical = True Then
                    
                    If OptionButtonMouseOver(0) = False And PropStyle = vbButtonGraphical Then
                        If OptionButtonEnabledVisualStyles = True And PropVisualStyles = True Then
                            OptionButtonMouseOver(0) = True
                            InvalidateRect hWnd, ByVal 0&, 0
                        End If
                    End If
                    
                    #End If
                    
                    If OptionButtonMouseOver(1) = False And PropMouseTrack = True Then
                        OptionButtonMouseOver(1) = True
                        RaiseEvent MouseEnter
                    End If
                    If OptionButtonMouseOver(0) = True Or OptionButtonMouseOver(1) = True Then Call ComCtlsRequestMouseLeave(hWnd)
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
        
        #If ImplementThemedGraphical = True Then
        
        If OptionButtonMouseOver(0) = True Then
            OptionButtonMouseOver(0) = False
            InvalidateRect hWnd, ByVal 0&, 0
        End If
        
        #End If
        
        If OptionButtonMouseOver(1) = True Then
            OptionButtonMouseOver(1) = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        If lParam = OptionButtonHandle Then
            Select Case HiWord(wParam)
                Case BN_CLICKED
                    If PropValue = False Then Me.Value = True
                Case BN_DOUBLECLICKED
                    RaiseEvent DblClick
            End Select
        End If
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = OptionButtonHandle Then
            Select Case NM.Code
                Case BCN_HOTITEMCHANGE
                    Dim NMBCHI As NMBCHOTITEM
                    CopyMemory NMBCHI, ByVal lParam, LenB(NMBCHI)
                    With NMBCHI
                    If (.dwFlags And HICF_MOUSE) = HICF_MOUSE Then
                        If (.dwFlags And HICF_ENTERING) = HICF_ENTERING Or (.dwFlags And HICF_LEAVING) = HICF_LEAVING Then RaiseEvent HotChanged
                    End If
                    End With
            End Select
        End If
    Case WM_CTLCOLORSTATIC, WM_CTLCOLORBTN
        WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
        If PropTransparent = True Then
            SetBkMode wParam, 1
            Dim hDCBmp As Long
            Dim hBmp As Long, hBmpOld As Long
            With UserControl
            If OptionButtonTransparentBrush = 0 Then
                hDCBmp = CreateCompatibleDC(wParam)
                If hDCBmp <> 0 Then
                    hBmp = CreateCompatibleBitmap(wParam, .ScaleWidth, .ScaleHeight)
                    If hBmp <> 0 Then
                        hBmpOld = SelectObject(hDCBmp, hBmp)
                        Dim WndRect As RECT, P As POINTAPI
                        GetWindowRect .hWnd, WndRect
                        MapWindowPoints HWND_DESKTOP, .ContainerHwnd, WndRect, 2
                        P.X = WndRect.Left
                        P.Y = WndRect.Top
                        SetViewportOrgEx hDCBmp, -P.X, -P.Y, P
                        SendMessage .ContainerHwnd, WM_PAINT, hDCBmp, ByVal 0&
                        SetViewportOrgEx hDCBmp, P.X, P.Y, P
                        OptionButtonTransparentBrush = CreatePatternBrush(hBmp)
                        SelectObject hDCBmp, hBmpOld
                        DeleteObject hBmp
                    End If
                    DeleteDC hDCBmp
                End If
            End If
            End With
            If OptionButtonTransparentBrush <> 0 Then WindowProcUserControl = OptionButtonTransparentBrush
        End If
        Exit Function
    Case WM_DRAWITEM
        Dim DIS As DRAWITEMSTRUCT
        CopyMemory DIS, ByVal lParam, LenB(DIS)
        If DIS.CtlType = ODT_BUTTON And DIS.hWndItem = OptionButtonHandle Then
            If (DIS.ItemAction And ODA_FOCUS) = ODA_FOCUS Then
                If PropValue = False Then Me.Value = True
            End If
            If PropStyle = vbButtonGraphical Then
                Dim Brush As Long, Text As String, TextRect As RECT
                Brush = CreateSolidBrush(WinColor(Me.BackColor))
                Text = Me.Caption
                Dim ButtonPicture As IPictureDisp, DisabledPictureAvailable As Boolean
                If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                    Set ButtonPicture = CoalescePicture(PropDisabledPicture, PropPicture)
                    If Not PropDisabledPicture Is Nothing Then
                        If PropDisabledPicture.Handle <> 0 Then DisabledPictureAvailable = True
                    End If
                ElseIf (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Or PropValue = True Then
                    Set ButtonPicture = CoalescePicture(PropDownPicture, PropPicture)
                Else
                    Set ButtonPicture = PropPicture
                End If
                If Not ButtonPicture Is Nothing Then
                    If ButtonPicture.Handle = 0 Then Set ButtonPicture = Nothing
                End If
                Dim hDCBmp2 As Long
                Dim hBmp2 As Long, hBmpOld2 As Long
                
                #If ImplementThemedGraphical = True Then
                
                Dim Theme As Long
                If OptionButtonEnabledVisualStyles = True And PropVisualStyles = True Then Theme = OpenThemeData(OptionButtonHandle, StrPtr("Button"))
                If Theme <> 0 Then
                    Dim ButtonPart As Long, ButtonState As Long
                    ButtonPart = BP_PUSHBUTTON
                    If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                        ButtonState = PBS_DISABLED
                    ElseIf (OptionButtonMouseOver(0) = True And PropValue = False) And Not (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Then
                        ButtonState = PBS_HOT
                    ElseIf (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Or PropValue = True Then
                        ButtonState = PBS_PRESSED
                    ElseIf (DIS.ItemState And ODS_FOCUS) = ODS_FOCUS Then
                        ButtonState = PBS_DEFAULTED
                    Else
                        ButtonState = PBS_NORMAL
                    End If
                    Dim RgnClip As Long
                    GetThemeBackgroundRegion Theme, DIS.hDC, ButtonPart, ButtonState, DIS.RCItem, RgnClip
                    ExtSelectClipRgn DIS.hDC, RgnClip, RGN_DIFF
                    If PropValue = True Then
                        If OptionButtonOwnerDrawCheckedBrush = 0 Then
                            hDCBmp2 = CreateCompatibleDC(DIS.hDC)
                            If hDCBmp2 <> 0 Then
                                hBmp2 = CreateCompatibleBitmap(DIS.hDC, 2, 2)
                                If hBmp2 <> 0 Then
                                    hBmpOld2 = SelectObject(hDCBmp2, hBmp2)
                                    SetPixel hDCBmp2, 0, 0, vbWhite
                                    SetPixel hDCBmp2, 1, 1, vbWhite
                                    SetPixel hDCBmp2, 0, 1, WinColor(UserControl.BackColor)
                                    SetPixel hDCBmp2, 1, 0, WinColor(UserControl.BackColor)
                                    OptionButtonOwnerDrawCheckedBrush = CreatePatternBrush(hBmp2)
                                    SelectObject hDCBmp2, hBmpOld2
                                    DeleteObject hBmp2
                                End If
                                DeleteDC hDCBmp2
                            End If
                        End If
                        If OptionButtonOwnerDrawCheckedBrush <> 0 Then FillRect DIS.hDC, DIS.RCItem, OptionButtonOwnerDrawCheckedBrush
                    Else
                        FillRect DIS.hDC, DIS.RCItem, Brush
                    End If
                    If IsThemeBackgroundPartiallyTransparent(Theme, ButtonPart, ButtonState) <> 0 Then DrawThemeParentBackground DIS.hWndItem, DIS.hDC, DIS.RCItem
                    ExtSelectClipRgn DIS.hDC, 0, RGN_COPY
                    DeleteObject RgnClip
                    DrawThemeBackground Theme, DIS.hDC, ButtonPart, ButtonState, DIS.RCItem, DIS.RCItem
                    GetThemeBackgroundContentRect Theme, DIS.hDC, ButtonPart, ButtonState, DIS.RCItem, DIS.RCItem
                    If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                        SetTextColor DIS.hDC, WinColor(vbGrayText)
                    Else
                        SetTextColor DIS.hDC, WinColor(Me.ForeColor)
                    End If
                    If (DIS.ItemState And ODS_FOCUS) = ODS_FOCUS Then
                        If Not (DIS.ItemState And ODS_NOFOCUSRECT) = ODS_NOFOCUSRECT Then DrawFocusRect DIS.hDC, DIS.RCItem
                    End If
                    If Not Text = vbNullString Then
                        LSet TextRect = DIS.RCItem
                        DrawText DIS.hDC, StrPtr(Text), -1, TextRect, DT_CALCRECT Or DT_WORDBREAK Or CLng(IIf((DIS.ItemState And ODS_NOACCEL) = ODS_NOACCEL, DT_HIDEPREFIX, 0))
                        TextRect.Left = DIS.RCItem.Left
                        TextRect.Right = DIS.RCItem.Right
                        If ButtonPicture Is Nothing Then
                            TextRect.Top = ((DIS.RCItem.Bottom - TextRect.Bottom) / 2) + (3 * PixelsPerDIP_Y())
                            TextRect.Bottom = TextRect.Top + TextRect.Bottom
                        Else
                            TextRect.Top = (DIS.RCItem.Bottom - TextRect.Bottom) + (1 * PixelsPerDIP_Y())
                            TextRect.Bottom = DIS.RCItem.Bottom
                        End If
                        If ComCtlsSupportLevel() >= 2 Then
                            Dim DTTO As DTTOPTS
                            DTTO.dwSize = LenB(DTTO)
                            DTTO.dwFlags = DTT_TEXTCOLOR
                            If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                                DTTO.crText = WinColor(Me.ForeColor)
                            Else
                                DTTO.crText = WinColor(vbGrayText)
                            End If
                            DrawThemeTextEx Theme, DIS.hDC, ButtonPart, ButtonState, StrPtr(Text), -1, DT_CENTER Or DT_WORDBREAK Or CLng(IIf((DIS.ItemState And ODS_NOACCEL) = ODS_NOACCEL, DT_HIDEPREFIX, 0)), TextRect, DTTO
                        Else
                            DrawThemeText Theme, DIS.hDC, ButtonPart, ButtonState, StrPtr(Text), -1, DT_CENTER Or DT_WORDBREAK Or CLng(IIf((DIS.ItemState And ODS_NOACCEL) = ODS_NOACCEL, DT_HIDEPREFIX, 0)), 0, TextRect
                        End If
                        DIS.RCItem.Bottom = TextRect.Top
                        DIS.RCItem.Left = TextRect.Left
                    End If
                    CloseThemeData Theme
                Else
                    Dim Flags As Long
                    Flags = DFCS_BUTTONPUSH
                    If (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Then Flags = Flags Or DFCS_PUSHED
                    If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then Flags = Flags Or DFCS_INACTIVE
                    If Me.Appearance = CCAppearanceFlat Then Flags = Flags Or DFCS_FLAT
                    If PropValue = True Then Flags = Flags Or DFCS_CHECKED
                    DrawFrameControl DIS.hDC, DIS.RCItem, DFC_BUTTON, Flags Or DFCS_ADJUSTRECT
                    If PropValue = True Then
                        If OptionButtonOwnerDrawCheckedBrush = 0 Then
                            hDCBmp2 = CreateCompatibleDC(DIS.hDC)
                            If hDCBmp2 <> 0 Then
                                hBmp2 = CreateCompatibleBitmap(DIS.hDC, 2, 2)
                                If hBmp2 <> 0 Then
                                    hBmpOld2 = SelectObject(hDCBmp2, hBmp2)
                                    SetPixel hDCBmp2, 0, 0, vbWhite
                                    SetPixel hDCBmp2, 1, 1, vbWhite
                                    SetPixel hDCBmp2, 0, 1, WinColor(UserControl.BackColor)
                                    SetPixel hDCBmp2, 1, 0, WinColor(UserControl.BackColor)
                                    OptionButtonOwnerDrawCheckedBrush = CreatePatternBrush(hBmp2)
                                    SelectObject hDCBmp2, hBmpOld2
                                    DeleteObject hBmp2
                                End If
                                DeleteDC hDCBmp2
                            End If
                        End If
                        If OptionButtonOwnerDrawCheckedBrush <> 0 Then FillRect DIS.hDC, DIS.RCItem, OptionButtonOwnerDrawCheckedBrush
                    Else
                        FillRect DIS.hDC, DIS.RCItem, Brush
                    End If
                    If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                        SetTextColor DIS.hDC, WinColor(vbGrayText)
                    Else
                        SetTextColor DIS.hDC, WinColor(Me.ForeColor)
                    End If
                    Call OffsetRect(DIS.RCItem, 1, 1, -1, -1)
                    If (DIS.ItemState And ODS_FOCUS) = ODS_FOCUS Then
                        If Not (DIS.ItemState And ODS_NOFOCUSRECT) = ODS_NOFOCUSRECT Then DrawFocusRect DIS.hDC, DIS.RCItem
                    End If
                    If Not Text = vbNullString Then
                        Dim OldBkMode As Long
                        OldBkMode = SetBkMode(DIS.hDC, 1)
                        LSet TextRect = DIS.RCItem
                        DrawText DIS.hDC, StrPtr(Text), -1, TextRect, DT_CALCRECT Or DT_WORDBREAK Or CLng(IIf((DIS.ItemState And ODS_NOACCEL) = ODS_NOACCEL, DT_HIDEPREFIX, 0))
                        TextRect.Left = DIS.RCItem.Left
                        TextRect.Right = DIS.RCItem.Right
                        If ButtonPicture Is Nothing Then
                            TextRect.Top = ((DIS.RCItem.Bottom - TextRect.Bottom) / 2) + (3 * PixelsPerDIP_Y())
                            TextRect.Bottom = TextRect.Top + TextRect.Bottom
                        Else
                            TextRect.Top = (DIS.RCItem.Bottom - TextRect.Bottom) + (1 * PixelsPerDIP_Y())
                            TextRect.Bottom = DIS.RCItem.Bottom
                        End If
                        If (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Or PropValue = True Then Call OffsetRect(TextRect, 1, 1, 1, 1)
                        DrawText DIS.hDC, StrPtr(Text), -1, TextRect, DT_CENTER Or DT_WORDBREAK Or CLng(IIf((DIS.ItemState And ODS_NOACCEL) = ODS_NOACCEL, DT_HIDEPREFIX, 0))
                        DIS.RCItem.Bottom = TextRect.Top
                        DIS.RCItem.Left = TextRect.Left
                        SetBkMode DIS.hDC, OldBkMode
                    End If
                End If
                
                #Else
                
                Dim Flags As Long
                Flags = DFCS_BUTTONPUSH
                If (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Then Flags = Flags Or DFCS_PUSHED
                If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then Flags = Flags Or DFCS_INACTIVE
                If Me.Appearance = CCAppearanceFlat Then Flags = Flags Or DFCS_FLAT
                If PropValue = True Then Flags = Flags Or DFCS_CHECKED
                DrawFrameControl DIS.hDC, DIS.RCItem, DFC_BUTTON, Flags Or DFCS_ADJUSTRECT
                If PropValue = True Then
                    If OptionButtonOwnerDrawCheckedBrush = 0 Then
                        hDCBmp2 = CreateCompatibleDC(DIS.hDC)
                        If hDCBmp2 <> 0 Then
                            hBmp2 = CreateCompatibleBitmap(DIS.hDC, 2, 2)
                            If hBmp2 <> 0 Then
                                hBmpOld2 = SelectObject(hDCBmp2, hBmp2)
                                SetPixel hDCBmp2, 0, 0, vbWhite
                                SetPixel hDCBmp2, 1, 1, vbWhite
                                SetPixel hDCBmp2, 0, 1, WinColor(UserControl.BackColor)
                                SetPixel hDCBmp2, 1, 0, WinColor(UserControl.BackColor)
                                OptionButtonOwnerDrawCheckedBrush = CreatePatternBrush(hBmp2)
                                SelectObject hDCBmp2, hBmpOld2
                                DeleteObject hBmp2
                            End If
                            DeleteDC hDCBmp2
                        End If
                    End If
                    If OptionButtonOwnerDrawCheckedBrush <> 0 Then FillRect DIS.hDC, DIS.RCItem, OptionButtonOwnerDrawCheckedBrush
                Else
                    FillRect DIS.hDC, DIS.RCItem, Brush
                End If
                If (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                    SetTextColor DIS.hDC, WinColor(vbGrayText)
                Else
                    SetTextColor DIS.hDC, WinColor(Me.ForeColor)
                End If
                Call OffsetRect(DIS.RCItem, 1, 1, -1, -1)
                If (DIS.ItemState And ODS_FOCUS) = ODS_FOCUS Then
                    If Not (DIS.ItemState And ODS_NOFOCUSRECT) = ODS_NOFOCUSRECT Then DrawFocusRect DIS.hDC, DIS.RCItem
                End If
                If Not Text = vbNullString Then
                    Dim OldBkMode As Long
                    OldBkMode = SetBkMode(DIS.hDC, 1)
                    LSet TextRect = DIS.RCItem
                    DrawText DIS.hDC, StrPtr(Text), -1, TextRect, DT_CALCRECT Or DT_WORDBREAK Or CLng(IIf((DIS.ItemState And ODS_NOACCEL) = ODS_NOACCEL, DT_HIDEPREFIX, 0))
                    TextRect.Left = DIS.RCItem.Left
                    TextRect.Right = DIS.RCItem.Right
                    If ButtonPicture Is Nothing Then
                        TextRect.Top = ((DIS.RCItem.Bottom - TextRect.Bottom) / 2) + (3 * PixelsPerDIP_Y())
                        TextRect.Bottom = TextRect.Top + TextRect.Bottom
                    Else
                        TextRect.Top = (DIS.RCItem.Bottom - TextRect.Bottom) + (1 * PixelsPerDIP_Y())
                        TextRect.Bottom = DIS.RCItem.Bottom
                    End If
                    If (DIS.ItemState And ODS_SELECTED) = ODS_SELECTED Or PropValue = True Then Call OffsetRect(TextRect, 1, 1, 1, 1)
                    DrawText DIS.hDC, StrPtr(Text), -1, TextRect, DT_CENTER Or DT_WORDBREAK Or CLng(IIf((DIS.ItemState And ODS_NOACCEL) = ODS_NOACCEL, DT_HIDEPREFIX, 0))
                    DIS.RCItem.Bottom = TextRect.Top
                    DIS.RCItem.Left = TextRect.Left
                    SetBkMode DIS.hDC, OldBkMode
                End If
                
                #End If
                
                If Not ButtonPicture Is Nothing Then
                    Dim CX As Long, CY As Long, X As Long, Y As Long
                    CX = CHimetricToPixel_X(ButtonPicture.Width)
                    CY = CHimetricToPixel_Y(ButtonPicture.Height)
                    X = DIS.RCItem.Left + ((DIS.RCItem.Right - DIS.RCItem.Left - CX) / 2)
                    Y = DIS.RCItem.Top + ((DIS.RCItem.Bottom - DIS.RCItem.Top - CY) / 2)
                    If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Or DisabledPictureAvailable = True Then
                        If ButtonPicture.Type = vbPicTypeBitmap And PropUseMaskColor = True Then
                            Dim hDC1 As Long, hBmpOld1 As Long
                            hDC1 = CreateCompatibleDC(DIS.hDC)
                            If hDC1 <> 0 Then
                                hBmpOld1 = SelectObject(hDC1, ButtonPicture.Handle)
                                TransparentBlt DIS.hDC, X, Y, CX, CY, hDC1, 0, 0, CX, CY, WinColor(PropMaskColor)
                                SelectObject hDC1, hBmpOld1
                                DeleteDC hDC1
                            End If
                        Else
                            Call RenderPicture(ButtonPicture, DIS.hDC, X, Y, CX, CY, OptionButtonPictureRenderFlag)
                        End If
                    Else
                        If ButtonPicture.Type = vbPicTypeIcon Then
                            DrawState DIS.hDC, 0, 0, ButtonPicture.Handle, 0, X, Y, CX, CY, DST_ICON Or DSS_DISABLED
                        Else
                            Dim hImage As Long
                            hImage = BitmapHandleFromPicture(ButtonPicture, vbWhite)
                            ' The DrawState API with DSS_DISABLED will draw white as transparent.
                            ' This will ensure GIF bitmaps or metafiles are better drawn.
                            DrawState DIS.hDC, 0, 0, hImage, 0, X, Y, CX, CY, DST_BITMAP Or DSS_DISABLED
                            DeleteObject hImage
                        End If
                    End If
                End If
                DeleteObject Brush
            Else
                With DIS
                RaiseEvent OwnerDraw(.ItemAction, .ItemState, .hDC, .RCItem.Left, .RCItem.Top, .RCItem.Right, .RCItem.Bottom)
                End With
            End If
            WindowProcUserControl = 1
            Exit Function
        End If
    Case UM_CHECKVALUE
        ' It is necessary to wait after all controls are initalized.
        ' If the property value is still valid here then notify the container so that option groups behave correctly.
        If PropValue = True Then UserControl.PropertyChanged "Value"
        Exit Function
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI OptionButtonHandle
End Function

Private Function WindowProcUserControlDesignMode(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_CTLCOLORBTN, WM_DRAWITEM
        WindowProcUserControlDesignMode = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
WindowProcUserControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
