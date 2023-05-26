VERSION 5.00
Begin VB.UserControl CoolBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ControlContainer=   -1  'True
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "CoolBar.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "CoolBar.ctx":0059
   Begin VB.Timer TimerInitChilds 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "CoolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

#Const ImplementThemedReBarFix = True

#If False Then
Private CbrOrientationHorizontal, CbrOrientationVertical
Private CbrBandStyleNormal, CbrBandStyleFixedSize
Private CbrBandGripperNormal, CbrBandGripperAlways, CbrBandGripperNever
Private CbrHitResultNoWhere, CbrHitResultCaption, CbrHitResultClient, CbrHitResultGrabber, CbrHitResultChevron, CbrHitResultSplitter
#End If
Public Enum CbrOrientationConstants
CbrOrientationHorizontal = 0
CbrOrientationVertical = 1
End Enum
Public Enum CbrBandStyleConstants
CbrBandStyleNormal = 0
CbrBandStyleFixedSize = 1
End Enum
Public Enum CbrBandGripperConstants
CbrBandGripperNormal = 0
CbrBandGripperAlways = 1
CbrBandGripperNever = 2
End Enum
Public Enum CbrHitResultConstants
CbrHitResultNoWhere = 0
CbrHitResultCaption = 1
CbrHitResultClient = 2
CbrHitResultGrabber = 3
CbrHitResultChevron = 4
CbrHitResultSplitter = 5
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
Private Type REBARINFO
cbSize As Long
fMask As Long
hImageList As Long
End Type
Private Type REBARBANDINFO
cbSize As Long
fMask As Long
fStyle As Long
ForeColor As Long
BackColor As Long
lpText As Long
cch As Long
iImage As Long
hWndChild As Long
CXMinChild As Long
CYMinChild As Long
CX As Long
hBmpBack As Long
wID As Long
CYChild As Long
CYMaxChild As Long
CYIntegral As Long
CXIdeal As Long
lParam As Long
CXHeader As Long
End Type
Private Type REBARBANDINFO_V61
RBBI As REBARBANDINFO
RCChevronLocation As RECT
uChevronState As Long
End Type
Private Type RBHITTESTINFO
PT As POINTAPI
Flag As Long
uBand As Long
End Type
Private Type PAINTSTRUCT
hDC As Long
fErase As Long
RCPaint As RECT
fRestore As Long
fIncUpdate As Long
RGBReserved(0 To 31) As Byte
End Type
Private Type TOOLINFO
cbSize As Long
uFlags As Long
hWnd As Long
uId As Long
RC As RECT
hInst As Long
lpszText As Long
lParam As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Const CDDS_PREPAINT As Long = &H1
Private Const CDDS_POSTPAINT As Long = &H2
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM + 1)
Private Const CDRF_DODEFAULT As Long = &H0
Private Const CDRF_SKIPDEFAULT As Long = &H4
Private Const CDRF_NOTIFYPOSTPAINT As Long = &H10
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
Private Type NMREBAR
hdr As NMHDR
dwMask As Long
uBand As Long
fStyle As Long
wID As Long
lParam As Long
End Type
Private Type NMREBARCHILDSIZE
hdr As NMHDR
uBand As Long
wID As Long
RCChild As RECT
RCBand As RECT
End Type
Private Type NMREBARCHEVRON
hdr As NMHDR
uBand As Long
wID As Long
lParam As Long
RCChevron As RECT
lParamNM As Long
End Type
Private Type NMTTDISPINFO
hdr As NMHDR
lpszText As Long
szText(0 To ((80 * 2) - 1)) As Byte
hInst As Long
uFlags As Long
lParam As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Public Event HeightChanged(ByVal NewHeight As Single)
Attribute HeightChanged.VB_Description = "Occurrs when the cool bar control's height changes, if its orientation is horizontal. Occurrs when the cool bar control's width changes, if its orientation is vertical."
Public Event LayoutChanged()
Attribute LayoutChanged.VB_Description = "Occurs when the user changes the layout of the bands."
Public Event MinMax(ByRef Cancel As Boolean)
Attribute MinMax.VB_Description = "Occurs prior to maximizing or minimizing a band."
Public Event BandBeforeDrag(ByVal Band As CbrBand, ByRef Cancel As Boolean)
Attribute BandBeforeDrag.VB_Description = "Occurs when a drag operation has begun on one band."
Public Event BandAfterDrag(ByVal Band As CbrBand, ByVal NewPosition As Long)
Attribute BandAfterDrag.VB_Description = "Occurs when a drag operation has ended on one band."
Public Event BandChevronPushed(ByVal Band As CbrBand, ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single)
Attribute BandChevronPushed.VB_Description = "Occurs when a chevron button of a band is pushed."
Public Event BandMouseEnter(ByVal Band As CbrBand)
Attribute BandMouseEnter.VB_Description = "Occurs when the user moves the mouse into a band."
Public Event BandMouseLeave(ByVal Band As CbrBand)
Attribute BandMouseLeave.VB_Description = "Occurs when the user moves the mouse out of a band."
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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal fMode As Long) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, ByRef CX As Long, ByRef CY As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32" (ByVal hImageList As Long, ByVal i As Long, ByVal hDcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long

#If ImplementThemedReBarFix = True Then

Private Enum UxThemeReBarParts
RP_GRIPPER = 1
RP_GRIPPERVERT = 2
RP_BAND = 3
RP_CHEVRON = 4
RP_CHEVRONVERT = 5
End Enum
Private Enum UxThemeChevronStates
CHEVS_NORMAL = 1
CHEVS_HOT = 2
CHEVS_PRESSED = 3
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
Private Declare Function DrawThemeText Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawThemeTextEx Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByRef lpRect As RECT, ByRef lpOptions As DTTOPTS) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal Theme As Long) As Long

#End If

Private Const ICC_COOL_CLASSES As Long = &H400
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOZORDER As Long = &H4
Private Const GWL_STYLE As Long = (-16)
Private Const TA_RTLREADING As Long = &H100
Private Const ILD_TRANSPARENT As Long = 1
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_LEFT As Long = &H0
Private Const DT_TOP As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_VCENTER As Long = &H4
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_TOPMOST As Long = &H8
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_BORDER As Long = &H800000
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
Private Const SW_SHOW As Long = &H5
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
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
Private Const WM_SETFONT As Long = &H30
Private Const WM_GETFONT As Long = &H31
Private Const WM_SIZE As Long = &H5
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINT As Long = &H317, PRF_CLIENT As Long = &H4, PRF_ERASEBKGND As Long = &H8
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETVERSION As Long = (CCM_FIRST + 7)
Private Const WM_USER As Long = &H400
Private Const RB_INSERTBANDA As Long = (WM_USER + 1)
Private Const RB_INSERTBANDW As Long = (WM_USER + 10)
Private Const RB_INSERTBAND As Long = RB_INSERTBANDW
Private Const RB_DELETEBAND As Long = (WM_USER + 2)
Private Const RB_GETBARINFO As Long = (WM_USER + 3)
Private Const RB_SETBARINFO As Long = (WM_USER + 4)
Private Const RB_SETBANDINFOA As Long = (WM_USER + 6)
Private Const RB_SETBANDINFOW As Long = (WM_USER + 11)
Private Const RB_SETBANDINFO As Long = RB_SETBANDINFOW
Private Const RB_SETPARENT As Long = (WM_USER + 7)
Private Const RB_HITTEST As Long = (WM_USER + 8)
Private Const RB_GETRECT As Long = (WM_USER + 9)
Private Const RB_GETBANDCOUNT As Long = (WM_USER + 12)
Private Const RB_GETROWCOUNT As Long = (WM_USER + 13)
Private Const RB_GETROWHEIGHT As Long = (WM_USER + 14)
Private Const RB_IDTOINDEX As Long = (WM_USER + 16)
Private Const RB_SETBKCOLOR As Long = (WM_USER + 19)
Private Const RB_GETBKCOLOR As Long = (WM_USER + 20)
Private Const RB_SETTEXTCOLOR As Long = (WM_USER + 21)
Private Const RB_GETTEXTCOLOR As Long = (WM_USER + 22)
Private Const RB_SIZETORECT As Long = (WM_USER + 23)
Private Const RB_BEGINDRAG As Long = (WM_USER + 24)
Private Const RB_ENDDRAG As Long = (WM_USER + 25)
Private Const RB_DRAGMOVE As Long = (WM_USER + 26)
Private Const RB_GETBARHEIGHT As Long = (WM_USER + 27)
Private Const RB_GETBANDINFOW As Long = (WM_USER + 28)
Private Const RB_GETBANDINFOA As Long = (WM_USER + 29)
Private Const RB_GETBANDINFO As Long = RB_GETBANDINFOW
Private Const RB_MINIMIZEBAND As Long = (WM_USER + 30)
Private Const RB_MAXIMIZEBAND As Long = (WM_USER + 31)
Private Const RB_GETBANDBORDERS As Long = (WM_USER + 34)
Private Const RB_SHOWBAND As Long = (WM_USER + 35)
Private Const RB_MOVEBAND As Long = (WM_USER + 39)
Private Const RB_GETBANDMARGINS As Long = (WM_USER + 40)
Private Const RB_PUSHCHEVRON As Long = (WM_USER + 43)
Private Const TTM_POP As Long = (WM_USER + 28)
Private Const TTM_ADDTOOLA As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW As Long = (WM_USER + 50)
Private Const TTM_ADDTOOL As Long = TTM_ADDTOOLW
Private Const TTM_NEWTOOLRECTA As Long = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW As Long = (WM_USER + 52)
Private Const TTM_NEWTOOLRECT As Long = TTM_NEWTOOLRECTW
Private Const LPSTR_TEXTCALLBACK As Long = (-1)
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTF_PARSELINKS As Long = &H1000
Private Const TTF_RTLREADING As Long = &H4
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFOA As Long = (TTN_FIRST - 0)
Private Const TTN_GETDISPINFOW As Long = (TTN_FIRST - 10)
Private Const TTN_GETDISPINFO As Long = TTN_GETDISPINFOW
Private Const RBBS_BREAK As Long = &H1
Private Const RBBS_FIXEDSIZE As Long = &H2
Private Const RBBS_CHILDEDGE As Long = &H4
Private Const RBBS_HIDDEN As Long = &H8
Private Const RBBS_NOVERT As Long = &H10 ' Malfunction
Private Const RBBS_FIXEDBMP As Long = &H20
Private Const RBBS_VARIABLEHEIGHT As Long = &H40
Private Const RBBS_GRIPPERALWAYS As Long = &H80
Private Const RBBS_NOGRIPPER As Long = &H100
Private Const RBBS_USECHEVRON As Long = &H200
Private Const RBBS_HIDETITLE As Long = &H400
Private Const RBBS_TOPALIGN As Long = &H800 ' Malfunction
Private Const RBBIM_STYLE As Long = &H1
Private Const RBBIM_COLORS As Long = &H2
Private Const RBBIM_TEXT As Long = &H4
Private Const RBBIM_IMAGE As Long = &H8
Private Const RBBIM_CHILD As Long = &H10
Private Const RBBIM_CHILDSIZE As Long = &H20
Private Const RBBIM_SIZE As Long = &H40
Private Const RBBIM_BACKGROUND As Long = &H80
Private Const RBBIM_ID As Long = &H100
Private Const RBBIM_IDEALSIZE As Long = &H200
Private Const RBBIM_LPARAM As Long = &H400
Private Const RBBIM_HEADERSIZE As Long = &H800
Private Const RBBIM_CHEVRONLOCATION As Long = &H1000
Private Const RBBIM_CHEVRONSTATE As Long = &H2000
Private Const RBHT_NOWHERE As Long = &H1
Private Const RBHT_CAPTION As Long = &H2
Private Const RBHT_CLIENT As Long = &H3
Private Const RBHT_GRABBER As Long = &H4
Private Const RBHT_CHEVRON As Long = &H8
Private Const RBHT_SPLITTER As Long = &H10
Private Const CCS_VERT As Long = &H80
Private Const CCS_NORESIZE As Long = &H4
Private Const CCS_NODIVIDER As Long = &H40
Private Const RBIM_IMAGELIST As Long = &H1
Private Const NM_FIRST As Long = 0
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Private Const RBN_FIRST As Long = (-831)
Private Const RBN_HEIGHTCHANGE As Long = (RBN_FIRST - 0)
Private Const RBN_LAYOUTCHANGED As Long = (RBN_FIRST - 2)
Private Const RBN_BEGINDRAG As Long = (RBN_FIRST - 4)
Private Const RBN_ENDDRAG As Long = (RBN_FIRST - 5)
Private Const RBN_DELETINGBAND As Long = (RBN_FIRST - 6)
Private Const RBN_DELETEDBAND As Long = (RBN_FIRST - 7)
Private Const RBN_CHILDSIZE As Long = (RBN_FIRST - 8)
Private Const RBN_CHEVRONPUSHED As Long = (RBN_FIRST - 10)
Private Const RBN_MINMAX As Long = (RBN_FIRST - 21)
Private Const RBNM_ID As Long = &H1
Private Const RBNM_STYLE As Long = &H2
Private Const RBNM_LPARAM As Long = &H4
Private Const RBS_TOOLTIPS As Long = &H100 ' Unsupported
Private Const RBS_VARHEIGHT As Long = &H200
Private Const RBS_BANDBORDERS As Long = &H400
Private Const RBS_FIXEDORDER As Long = &H800
Private Const RBS_VERTICALGRIPPER As Long = &H4000
Private Const RBS_DBLCLKTOGGLE As Long = &H8000&
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IPerPropertyBrowsingVB
Private Type InitBandStruct
Caption As String
Key As String
Tag As String
ChildName As String
Style As CbrBandStyleConstants
Image As Variant
ImageIndex As Long
Width As Long
MinWidth As Long
MinHeight As Long
IdealWidth As Long
Gripper As CbrBandGripperConstants
ToolTipText As String
UseCoolBarPicture As Boolean
Picture As IPictureDisp
UseCoolBarColors As Boolean
BackColor As OLE_COLOR
ForeColor As OLE_COLOR
NewRow As Boolean
Visible As Boolean
ChildEdge As Boolean
UseChevron As Boolean
HideCaption As Boolean
FixedBackground As Boolean
End Type
Private CoolBarHandle As Long, CoolBarToolTipHandle As Long
Private CoolBarFontHandle As Long
Private CoolBarIsClick As Boolean
Private CoolBarMouseOver As Boolean, CoolBarMouseOverIndex As Long
Private CoolBarDesignMode As Boolean
Private CoolBarToolTipIndex As Long
Private CoolBarDoubleBufferEraseBkgDC As Long
Private CoolBarAlignable As Boolean
Private CoolBarTheme As Long
Private CoolBarImageListObjectPointer As Long
Private DispIDMousePointer As Long
Private DispIDBorderStyle As Long
Private DispIDImageList As Long, ImageListArray() As String
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropBands As CbrBands
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropImageListName As String, PropImageListInit As Boolean
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR
Private PropBorderStyle As Integer
Private PropOrientation As CbrOrientationConstants
Private PropBandBorders As Boolean
Private PropFixedOrder As Boolean
Private PropVariantHeight As Boolean
Private PropPicture As IPictureDisp
Private PropDblClickToggle As Boolean
Private PropVerticalGripper As Boolean
Private PropShowTips As Boolean
Private PropDoubleBuffer As Boolean

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetDisplayStringMousePointer(PropMousePointer, DisplayName)
    Handled = True
ElseIf DispID = DispIDBorderStyle Then
    Select Case PropBorderStyle
        Case vbBSNone: DisplayName = vbBSNone & " - None"
        Case vbFixedSingle: DisplayName = vbFixedSingle & " - Fixed Single"
    End Select
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
ElseIf DispID = DispIDBorderStyle Then
    ReDim StringsOut(0 To (1 + 1)) As String
    ReDim CookiesOut(0 To (1 + 1)) As Long
    StringsOut(0) = vbBSNone & " - None": CookiesOut(0) = vbBSNone
    StringsOut(1) = vbFixedSingle & " - Fixed Single": CookiesOut(1) = vbFixedSingle
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
ElseIf DispID = DispIDBorderStyle Then
    Value = Cookie
    Handled = True
ElseIf DispID = DispIDImageList Then
    If Cookie < UBound(ImageListArray()) Then Value = ImageListArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_COOL_CLASSES)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim ImageListArray(0) As String
CoolBarToolTipIndex = -1
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDBorderStyle = 0 Then DispIDBorderStyle = GetDispID(Me, "BorderStyle")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then CoolBarAlignable = False Else CoolBarAlignable = True
CoolBarDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = "(None)"
PropBackColor = vbButtonFace
PropForeColor = vbButtonText
PropBorderStyle = vbFixedSingle
PropOrientation = CbrOrientationHorizontal
PropBandBorders = True
PropFixedOrder = False
PropVariantHeight = True
Set PropPicture = Nothing
PropDblClickToggle = False
PropVerticalGripper = False
PropShowTips = False
PropDoubleBuffer = True
Call CreateCoolBar
Me.Bands.Add().Width = UserControl.ScaleX((192 * PixelsPerDIP_X()), vbPixels, vbContainerSize)
Me.Bands.Add(NewRow:=True).Width = UserControl.ScaleX((96 * PixelsPerDIP_X()), vbPixels, vbContainerSize)
Me.Bands.Add().Width = UserControl.ScaleX((96 * PixelsPerDIP_X()), vbPixels, vbContainerSize)
End Sub

Private Sub UserControl_Paint()
If CoolBarDesignMode = True Then RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDBorderStyle = 0 Then DispIDBorderStyle = GetDispID(Me, "BorderStyle")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then CoolBarAlignable = False Else CoolBarAlignable = True
CoolBarDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = .ReadProperty("ImageList", "(None)")
PropBackColor = .ReadProperty("BackColor", vbButtonFace)
PropForeColor = .ReadProperty("ForeColor", vbButtonText)
PropBorderStyle = .ReadProperty("BorderStyle", vbFixedSingle)
PropOrientation = .ReadProperty("Orientation", CbrOrientationHorizontal)
PropBandBorders = .ReadProperty("BandBorders", True)
PropFixedOrder = .ReadProperty("FixedOrder", False)
PropVariantHeight = .ReadProperty("VariantHeight", True)
Set PropPicture = .ReadProperty("Picture", Nothing)
PropDblClickToggle = .ReadProperty("DblClickToggle", False)
PropVerticalGripper = .ReadProperty("VerticalGripper", False)
PropShowTips = .ReadProperty("ShowTips", False)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
End With
Dim InitChilds As Boolean
With New PropertyBag
On Error Resume Next
.Contents = PropBag.ReadProperty("InitBands", 0)
On Error GoTo 0
Dim InitBandsCount As Long, i As Long
Dim InitBands() As InitBandStruct
InitBandsCount = .ReadProperty("InitBandsCount", 0)
If InitBandsCount > 0 Then
    ReDim InitBands(1 To InitBandsCount) As InitBandStruct
    Dim VarValue As Variant
    For i = 1 To InitBandsCount
        InitBands(i).Caption = VarToStr(.ReadProperty("InitBandsCaption" & CStr(i), vbNullString))
        InitBands(i).Key = VarToStr(.ReadProperty("InitBandsKey" & CStr(i), vbNullString))
        InitBands(i).Tag = VarToStr(.ReadProperty("InitBandsTag" & CStr(i), vbNullString))
        InitBands(i).ChildName = VarToStr(.ReadProperty("InitBandsChildName" & CStr(i), vbNullString))
        If Not InitBands(i).ChildName = vbNullString Then InitChilds = True
        InitBands(i).Style = .ReadProperty("InitBandsStyle" & CStr(i), CbrBandStyleNormal)
        VarValue = .ReadProperty("InitBandsImage" & CStr(i), 0)
        If VarType(VarValue) = vbArray + vbByte Then
            InitBands(i).Image = VarToStr(VarValue)
            InitBands(i).ImageIndex = .ReadProperty("InitBandsImageIndex" & CStr(i), 0)
        Else
            InitBands(i).Image = VarValue
            InitBands(i).ImageIndex = CLng(VarValue)
        End If
        InitBands(i).Width = (.ReadProperty("InitBandsWidth" & CStr(i), 0) * PixelsPerDIP_X())
        InitBands(i).MinWidth = (.ReadProperty("InitBandsMinWidth" & CStr(i), 0) * PixelsPerDIP_X())
        InitBands(i).MinHeight = (.ReadProperty("InitBandsMinHeight" & CStr(i), 0) * PixelsPerDIP_Y())
        InitBands(i).IdealWidth = (.ReadProperty("InitBandsIdealWidth" & CStr(i), 0) * PixelsPerDIP_X())
        InitBands(i).Gripper = .ReadProperty("InitBandsGripper" & CStr(i), CbrBandGripperNormal)
        InitBands(i).ToolTipText = VarToStr(.ReadProperty("InitBandsToolTipText" & CStr(i), vbNullString))
        InitBands(i).UseCoolBarPicture = .ReadProperty("InitBandsUseCoolBarPicture" & CStr(i), True)
        Set InitBands(i).Picture = .ReadProperty("InitBandsPicture" & CStr(i), Nothing)
        InitBands(i).UseCoolBarColors = .ReadProperty("InitBandsUseCoolBarColors" & CStr(i), True)
        InitBands(i).BackColor = .ReadProperty("InitBandsBackColor" & CStr(i), vbButtonFace)
        InitBands(i).ForeColor = .ReadProperty("InitBandsForeColor" & CStr(i), vbButtonText)
        InitBands(i).NewRow = .ReadProperty("InitBandsNewRow" & CStr(i), False)
        InitBands(i).Visible = .ReadProperty("InitBandsVisible" & CStr(i), True)
        InitBands(i).ChildEdge = .ReadProperty("InitBandsChildEdge" & CStr(i), False)
        InitBands(i).UseChevron = .ReadProperty("InitBandsUseChevron" & CStr(i), False)
        InitBands(i).HideCaption = .ReadProperty("InitBandsHideCaption" & CStr(i), False)
        InitBands(i).FixedBackground = .ReadProperty("InitBandsFixedBackground" & CStr(i), True)
    Next i
End If
End With
Call CreateCoolBar
If InitBandsCount > 0 And CoolBarHandle <> 0 Then
    Dim ImageListInit As Boolean
    ImageListInit = PropImageListInit
    PropImageListInit = True
    For i = 1 To InitBandsCount
        With Me.Bands.Add(i, InitBands(i).Key, InitBands(i).Caption, InitBands(i).ImageIndex, InitBands(i).NewRow, Nothing, InitBands(i).Visible)
        .FInit Me, InitBands(i).Key, InitBands(i).ChildName, InitBands(i).Image, InitBands(i).ImageIndex
        .Tag = InitBands(i).Tag
        .Style = InitBands(i).Style
        If InitBands(i).Style <> CbrBandStyleFixedSize Then
            .Width = UserControl.ScaleX(InitBands(i).Width, vbPixels, vbContainerSize)
        End If
        .MinWidth = UserControl.ScaleX(InitBands(i).MinWidth, vbPixels, vbContainerSize)
        .MinHeight = UserControl.ScaleY(InitBands(i).MinHeight, vbPixels, vbContainerSize)
        .IdealWidth = UserControl.ScaleX(InitBands(i).IdealWidth, vbPixels, vbContainerSize)
        .Gripper = InitBands(i).Gripper
        .ToolTipText = InitBands(i).ToolTipText
        .UseCoolBarPicture = InitBands(i).UseCoolBarPicture
        Set .Picture = InitBands(i).Picture
        .UseCoolBarColors = InitBands(i).UseCoolBarColors
        .BackColor = InitBands(i).BackColor
        .ForeColor = InitBands(i).ForeColor
        .ChildEdge = InitBands(i).ChildEdge
        .UseChevron = InitBands(i).UseChevron
        .HideCaption = InitBands(i).HideCaption
        .FixedBackground = InitBands(i).FixedBackground
        End With
    Next i
    PropImageListInit = ImageListInit
End If
If Not PropImageListName = "(None)" Then TimerImageList.Enabled = True
If InitChilds = True Then TimerInitChilds.Enabled = True
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
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "ImageList", PropImageListName, "(None)"
.WriteProperty "BackColor", PropBackColor, vbButtonFace
.WriteProperty "ForeColor", PropForeColor, vbButtonText
.WriteProperty "BorderStyle", PropBorderStyle, vbFixedSingle
.WriteProperty "Orientation", PropOrientation, CbrOrientationHorizontal
.WriteProperty "BandBorders", PropBandBorders, True
.WriteProperty "FixedOrder", PropFixedOrder, False
.WriteProperty "VariantHeight", PropVariantHeight, True
.WriteProperty "Picture", PropPicture, Nothing
.WriteProperty "DblClickToggle", PropDblClickToggle, False
.WriteProperty "VerticalGripper", PropVerticalGripper, False
.WriteProperty "ShowTips", PropShowTips, False
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
End With
Dim Count As Long
Count = Me.Bands.Count
With New PropertyBag
.WriteProperty "InitBandsCount", Count, 0
If Count > 0 Then
    Dim i As Long, VarValue As Variant
    For i = 1 To Count
        .WriteProperty "InitBandsCaption" & CStr(i), StrToVar(Me.Bands(i).Caption), vbNullString
        .WriteProperty "InitBandsKey" & CStr(i), StrToVar(Me.Bands(i).Key), vbNullString
        .WriteProperty "InitBandsTag" & CStr(i), StrToVar(Me.Bands(i).Tag), vbNullString
        .WriteProperty "InitBandsChildName" & CStr(i), StrToVar(ProperControlName(Me.Bands(i).Child)), vbNullString
        .WriteProperty "InitBandsStyle" & CStr(i), Me.Bands(i).Style, CbrBandStyleNormal
        VarValue = Me.Bands(i).Image
        If VarType(VarValue) = vbString Then
            .WriteProperty "InitBandsImage" & CStr(i), StrToVar(VarValue), 0
            .WriteProperty "InitBandsImageIndex" & CStr(i), Me.Bands(i).ImageIndex, 0
        Else
            .WriteProperty "InitBandsImage" & CStr(i), VarValue, 0
        End If
        .WriteProperty "InitBandsWidth" & CStr(i), (CLng(UserControl.ScaleX(Me.Bands(i).Width, vbContainerSize, vbPixels)) / PixelsPerDIP_X()), 0
        .WriteProperty "InitBandsMinWidth" & CStr(i), (CLng(UserControl.ScaleX(Me.Bands(i).MinWidth, vbContainerSize, vbPixels)) / PixelsPerDIP_X()), 0
        .WriteProperty "InitBandsMinHeight" & CStr(i), (CLng(UserControl.ScaleY(Me.Bands(i).MinHeight, vbContainerSize, vbPixels)) / PixelsPerDIP_Y()), 0
        .WriteProperty "InitBandsIdealWidth" & CStr(i), (CLng(UserControl.ScaleX(Me.Bands(i).IdealWidth, vbContainerSize, vbPixels)) / PixelsPerDIP_X()), 0
        .WriteProperty "InitBandsGripper" & CStr(i), Me.Bands(i).Gripper, CbrBandGripperNormal
        .WriteProperty "InitBandsToolTipText" & CStr(i), StrToVar(Me.Bands(i).ToolTipText), vbNullString
        .WriteProperty "InitBandsUseCoolBarPicture" & CStr(i), Me.Bands(i).UseCoolBarPicture, True
        .WriteProperty "InitBandsPicture" & CStr(i), Me.Bands(i).Picture, Nothing
        .WriteProperty "InitBandsUseCoolBarColors" & CStr(i), Me.Bands(i).UseCoolBarColors, True
        .WriteProperty "InitBandsBackColor" & CStr(i), Me.Bands(i).BackColor, vbButtonFace
        .WriteProperty "InitBandsForeColor" & CStr(i), Me.Bands(i).ForeColor, vbButtonText
        .WriteProperty "InitBandsNewRow" & CStr(i), Me.Bands(i).NewRow, False
        .WriteProperty "InitBandsVisible" & CStr(i), Me.Bands(i).Visible, True
        .WriteProperty "InitBandsChildEdge" & CStr(i), Me.Bands(i).ChildEdge, False
        .WriteProperty "InitBandsUseChevron" & CStr(i), Me.Bands(i).UseChevron, False
        .WriteProperty "InitBandsHideCaption" & CStr(i), Me.Bands(i).HideCaption, False
        .WriteProperty "InitBandsFixedBackground" & CStr(i), Me.Bands(i).FixedBackground, True
    Next i
End If
PropBag.WriteProperty "InitBands", .Contents, 0
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
Static LastHeight As Single, LastWidth As Single, LastAlign As Integer
Static PrevHeight As Long, PrevWidth As Long
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl.Extender
Dim Align As Integer
If CoolBarAlignable = True Then Align = .Align Else Align = vbAlignNone
Select Case Align
    Case LastAlign
    Case vbAlignNone
    Case vbAlignTop, vbAlignBottom
        Select Case LastAlign
            Case vbAlignLeft, vbAlignRight
                .Height = LastWidth
        End Select
    Case vbAlignLeft, vbAlignRight
        Select Case LastAlign
            Case vbAlignTop, vbAlignBottom
                .Width = LastHeight
        End Select
End Select
LastHeight = .Height
LastWidth = .Width
LastAlign = Align
End With
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If CoolBarHandle = 0 Then InProc = False: Exit Sub
Dim Count As Long, Size As SIZEAPI, WndRect As RECT, Rows As Long
Count = SendMessage(CoolBarHandle, RB_GETBANDCOUNT, 0, ByVal 0&)
GetWindowRect CoolBarHandle, WndRect
If Count > 0 Then
    Dim ClientRect As RECT
    GetClientRect CoolBarHandle, ClientRect
    Rows = SendMessage(CoolBarHandle, RB_GETROWCOUNT, 0, ByVal 0&)
    If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
        Size.CY = SendMessage(CoolBarHandle, RB_GETBARHEIGHT, 0, ByVal 0&) + ((WndRect.Bottom - WndRect.Top) - (ClientRect.Bottom - ClientRect.Top))
        Size.CX = UserControl.ScaleWidth
    Else
        Size.CX = SendMessage(CoolBarHandle, RB_GETBARHEIGHT, 0, ByVal 0&) + ((WndRect.Bottom - WndRect.Top) - (ClientRect.Bottom - ClientRect.Top))
        Size.CY = UserControl.ScaleHeight
    End If
Else
    Size.CY = UserControl.ScaleHeight
    Size.CX = UserControl.ScaleWidth
End If
Select Case Align
    Case vbAlignNone
        .Extender.Move .Extender.Left, .Extender.Top, .ScaleX(Size.CX, vbPixels, vbContainerSize), .ScaleY(Size.CY, vbPixels, vbContainerSize)
    Case vbAlignTop, vbAlignBottom
        .Extender.Height = .ScaleY(Size.CY, vbPixels, vbContainerSize)
    Case vbAlignLeft, vbAlignRight
        .Extender.Width = .ScaleX(Size.CX, vbPixels, vbContainerSize)
End Select
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
MoveWindow CoolBarHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
If Count > 0 Then
    If Rows <> SendMessage(CoolBarHandle, RB_GETROWCOUNT, 0, ByVal 0&) Then Call UserControl_Resize: Exit Sub
End If
With UserControl
If PrevHeight <> .ScaleHeight Or PrevWidth <> .ScaleWidth Then
    PrevHeight = .ScaleHeight
    PrevWidth = .ScaleWidth
    RaiseEvent Resize
End If
End With
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyCoolBar
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropImageListInit = False Then
    If Not PropImageListName = "(None)" Then Me.ImageList = PropImageListName
    PropImageListInit = True
End If
TimerImageList.Enabled = False
End Sub

Private Sub TimerInitChilds_Timer()
Dim Count As Long
Count = Me.Bands.Count
If Count > 0 Then
    Dim i As Long
    If CoolBarDesignMode = False Then
        For i = 1 To Count
            With Me.Bands(i)
            Set .Child = .Child
            End With
        Next i
    Else
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_CHILD
        For i = 0 To Count - 1
            .hWndChild = 0
            SendMessage CoolBarHandle, RB_SETBANDINFO, i, ByVal VarPtr(RBBI)
            .hWndChild = -1
            SendMessage CoolBarHandle, RB_SETBANDINFO, i, ByVal VarPtr(RBBI)
        Next i
        End With
        Call UserControl_Resize
    End If
End If
TimerInitChilds.Enabled = False
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

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
Attribute WhatsThisHelpID.VB_MemberFlags = "400"
WhatsThisHelpID = Extender.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal Value As Long)
Extender.WhatsThisHelpID = Value
End Property

Public Property Get Align() As Integer
Attribute Align.VB_Description = "Returns/sets a value that determines where an object is displayed on a form."
Attribute Align.VB_MemberFlags = "400"
Align = Extender.Align
End Property

Public Property Let Align(ByVal Value As Integer)
Extender.Align = Value
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

Public Sub ZOrder(Optional ByRef Position As Variant)
Attribute ZOrder.VB_Description = "Places a specified object at the front or back of the z-order within its graphical level."
If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
hWnd = CoolBarHandle
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
OldFontHandle = CoolBarFontHandle
CoolBarFontHandle = CreateGDIFontFromOLEFont(PropFont)
If CoolBarHandle <> 0 Then SendMessage CoolBarHandle, WM_SETFONT, CoolBarFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = CoolBarFontHandle
CoolBarFontHandle = CreateGDIFontFromOLEFont(PropFont)
If CoolBarHandle <> 0 Then SendMessage CoolBarHandle, WM_SETFONT, CoolBarFontHandle, ByVal 1&
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
If CoolBarHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles CoolBarHandle
    Else
        RemoveVisualStyles CoolBarHandle
    End If
    Call SetVisualStylesToolTip
    If PropBorderStyle <> vbBSNone Then
        Call ComCtlsChangeBorderStyle(CoolBarHandle, CCBorderStyleSingle)
    Else
        Call ComCtlsChangeBorderStyle(CoolBarHandle, CCBorderStyleNone)
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
If CoolBarHandle <> 0 Then EnableWindow CoolBarHandle, IIf(Value = True, 1, 0)
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
If CoolBarDesignMode = False Then Call RefreshMousePointer
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
        If CoolBarDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If CoolBarDesignMode = False Then Call RefreshMousePointer
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
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
If CoolBarDesignMode = False Then Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
If CoolBarHandle <> 0 Then
    Call ComCtlsSetRightToLeft(CoolBarHandle, dwMask)
    If CoolBarDesignMode = False Then
        Dim Band As CbrBand, Child As Object
        For Each Band In Me.Bands
            Set Child = Band.Child
            If Not Child Is Nothing Then Child.Move Child.Left
        Next Band
    Else
        Dim RBBI As REBARBANDINFO, i As Long
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_CHILD
        For i = 1 To SendMessage(CoolBarHandle, RB_GETBANDCOUNT, 0, ByVal 0&)
            .hWndChild = 0
            SendMessage CoolBarHandle, RB_SETBANDINFO, i - 1, ByVal VarPtr(RBBI)
            .hWndChild = -1
            SendMessage CoolBarHandle, RB_SETBANDINFO, i - 1, ByVal VarPtr(RBBI)
        Next i
        End With
    End If
End If
If CoolBarToolTipHandle <> 0 Then
    If PropRightToLeft = True Then
        If PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = WS_EX_RTLREADING
    Else
        dwMask = 0
    End If
    Call ComCtlsSetRightToLeft(CoolBarToolTipHandle, dwMask)
End If
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
If CoolBarDesignMode = False Then
    If PropImageListInit = False And CoolBarImageListObjectPointer = 0 Then
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
If CoolBarHandle <> 0 Then
    Dim RBI As REBARINFO
    RBI.cbSize = LenB(RBI)
    RBI.fMask = RBIM_IMAGELIST
    Dim Success As Boolean, Handle As Long
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> 0)
        End If
        If Success = True Then
            RBI.hImageList = Handle
            SendMessage CoolBarHandle, RB_SETBARINFO, 0, ByVal VarPtr(RBI)
            CoolBarImageListObjectPointer = ObjPtr(Value)
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
                        RBI.hImageList = Handle
                        SendMessage CoolBarHandle, RB_SETBARINFO, 0, ByVal VarPtr(RBI)
                        If CoolBarDesignMode = False Then CoolBarImageListObjectPointer = ObjPtr(ControlEnum)
                        PropImageListName = Value
                        Exit For
                    ElseIf CoolBarDesignMode = True Then
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
        SendMessage CoolBarHandle, RB_GETBARINFO, 0, ByVal VarPtr(RBI)
        If RBI.hImageList <> 0 Then
            RBI.hImageList = 0
            SendMessage CoolBarHandle, RB_SETBARINFO, 0, ByVal VarPtr(RBI)
        End If
        CoolBarImageListObjectPointer = 0
        PropImageListName = "(None)"
    ElseIf Handle = 0 Then
        SendMessage CoolBarHandle, RB_GETBARINFO, 0, ByVal VarPtr(RBI)
        If RBI.hImageList <> 0 Then
            RBI.hImageList = 0
            SendMessage CoolBarHandle, RB_SETBARINFO, 0, ByVal VarPtr(RBI)
        End If
    End If
    Me.Refresh
End If
UserControl.PropertyChanged "ImageList"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object. This property is ignored if the version of comctl32.dll is 6.0 or higher and the visual styles property is set to true."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If CoolBarHandle <> 0 Then SendMessage CoolBarHandle, RB_SETBKCOLOR, 0, ByVal WinColor(PropBackColor)
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If CoolBarHandle <> 0 Then
    SendMessage CoolBarHandle, RB_SETTEXTCOLOR, 0, ByVal WinColor(PropForeColor)
    Me.Refresh
End If
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_UserMemId = -504
BorderStyle = PropBorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As Integer)
Select Case Value
    Case vbBSNone, vbFixedSingle
        PropBorderStyle = Value
    Case Else
        Err.Raise 380
End Select
If CoolBarHandle <> 0 Then
    If PropBorderStyle <> vbBSNone Then
        Call ComCtlsChangeBorderStyle(CoolBarHandle, CCBorderStyleSingle)
    Else
        Call ComCtlsChangeBorderStyle(CoolBarHandle, CCBorderStyleNone)
    End If
    Call UserControl_Resize
End If
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get Orientation() As CbrOrientationConstants
Attribute Orientation.VB_Description = "Returns/sets the orientation."
Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As CbrOrientationConstants)
Dim DoSwap As Boolean, SwapWidth As Single, SwapHeight As Single
Select Case Value
    Case CbrOrientationHorizontal, CbrOrientationVertical
        Dim Align As Integer
        If CoolBarAlignable = True Then Align = Extender.Align Else Align = vbAlignNone
        If Align = vbAlignNone Then
            If PropOrientation = CbrOrientationHorizontal And Value = CbrOrientationVertical Then
                DoSwap = True
                SwapWidth = UserControl.Width
                SwapHeight = UserControl.Height
            ElseIf PropOrientation = CbrOrientationVertical And Value = CbrOrientationHorizontal Then
                DoSwap = True
                SwapWidth = UserControl.Width
                SwapHeight = UserControl.Height
            End If
        End If
        PropOrientation = Value
    Case Else
        Err.Raise 380
End Select
If CoolBarHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(CoolBarHandle, GWL_STYLE)
    Select Case PropOrientation
        Case CbrOrientationHorizontal
            If (dwStyle And CCS_VERT) = CCS_VERT Then dwStyle = dwStyle And Not CCS_VERT
        Case CbrOrientationVertical
            If Not (dwStyle And CCS_VERT) = CCS_VERT Then dwStyle = dwStyle Or CCS_VERT
    End Select
    SetWindowLong CoolBarHandle, GWL_STYLE, dwStyle
    Me.Refresh
    Call UserControl_Resize
    If DoSwap = True Then
        UserControl.Width = SwapHeight
        UserControl.Height = SwapWidth
    End If
End If
UserControl.PropertyChanged "Orientation"
End Property

Public Property Get BandBorders() As Boolean
Attribute BandBorders.VB_Description = "Returns/sets a value indicating whether the cool bar displays narrow lines to separate the bands."
BandBorders = PropBandBorders
End Property

Public Property Let BandBorders(ByVal Value As Boolean)
PropBandBorders = Value
If CoolBarHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(CoolBarHandle, GWL_STYLE)
    If PropBandBorders = True Then
        If Not (dwStyle And RBS_BANDBORDERS) = RBS_BANDBORDERS Then dwStyle = dwStyle Or RBS_BANDBORDERS
    Else
        If (dwStyle And RBS_BANDBORDERS) = RBS_BANDBORDERS Then dwStyle = dwStyle And Not RBS_BANDBORDERS
    End If
    SetWindowLong CoolBarHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "BandBorders"
End Property

Public Property Get FixedOrder() As Boolean
Attribute FixedOrder.VB_Description = "Returns/sets a value indicating whether the user is allowed to rearrange the order of the bands."
FixedOrder = PropFixedOrder
End Property

Public Property Let FixedOrder(ByVal Value As Boolean)
PropFixedOrder = Value
If CoolBarHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(CoolBarHandle, GWL_STYLE)
    If PropFixedOrder = True Then
        If Not (dwStyle And RBS_FIXEDORDER) = RBS_FIXEDORDER Then dwStyle = dwStyle Or RBS_FIXEDORDER
    Else
        If (dwStyle And RBS_FIXEDORDER) = RBS_FIXEDORDER Then dwStyle = dwStyle And Not RBS_FIXEDORDER
    End If
    SetWindowLong CoolBarHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "FixedOrder"
End Property

Public Property Get VariantHeight() As Boolean
Attribute VariantHeight.VB_Description = "Returns/sets a value indicating whether the cool bar allows bands to be displayed with different heights."
VariantHeight = PropVariantHeight
End Property

Public Property Let VariantHeight(ByVal Value As Boolean)
PropVariantHeight = Value
If CoolBarHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(CoolBarHandle, GWL_STYLE)
    If PropVariantHeight = True Then
        If Not (dwStyle And RBS_VARHEIGHT) = RBS_VARHEIGHT Then dwStyle = dwStyle Or RBS_VARHEIGHT
    Else
        If (dwStyle And RBS_VARHEIGHT) = RBS_VARHEIGHT Then dwStyle = dwStyle And Not RBS_VARHEIGHT
    End If
    SetWindowLong CoolBarHandle, GWL_STYLE, dwStyle
    If (dwStyle And CCS_VERT) = 0 Then
        SetWindowPos CoolBarHandle, 0, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight + 10, SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or &H20
    Else
        SetWindowPos CoolBarHandle, 0, 0, 0, UserControl.ScaleWidth + 10, UserControl.ScaleHeight, SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or &H20
    End If
    Me.Refresh
    Call UserControl_Resize
End If
UserControl.PropertyChanged "VariantHeight"
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets the background picture."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Set Me.Picture = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropPicture = Nothing
Else
    If Value.Type = vbPicTypeBitmap Or Value.Handle = 0 Then
        Set PropPicture = Value
    Else
        If CoolBarDesignMode = True Then
            MsgBox "Invalid picture", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 481
        End If
    End If
End If
If CoolBarHandle <> 0 Then
    Dim Count As Long
    Count = SendMessage(CoolBarHandle, RB_GETBANDCOUNT, 0, ByVal 0&)
    If Count > 0 Then
        Dim i As Long, hBmp As Long
        Dim RBBI As REBARBANDINFO, Band As CbrBand
        With RBBI
        .cbSize = LenB(RBBI)
        If PropPicture Is Nothing Then hBmp = 0 Else hBmp = PropPicture.Handle
        For i = 0 To Count - 1
            .fMask = RBBIM_LPARAM Or RBBIM_BACKGROUND
            SendMessage CoolBarHandle, RB_GETBANDINFO, i, ByVal VarPtr(RBBI)
            If .lParam <> 0 Then
                Set Band = PtrToObj(.lParam)
                If Band.UseCoolBarPicture = True Then
                    .fMask = RBBIM_BACKGROUND
                    .hBmpBack = hBmp
                    SendMessage CoolBarHandle, RB_SETBANDINFO, i, ByVal VarPtr(RBBI)
                End If
            End If
        Next i
        End With
    End If
End If
UserControl.PropertyChanged "Picture"
End Property

Public Property Get DblClickToggle() As Boolean
Attribute DblClickToggle.VB_Description = "Returns/sets a value that determines whether or not the cool bar will toggle its maximized or minimized state when the user double-clicks a band."
DblClickToggle = PropDblClickToggle
End Property

Public Property Let DblClickToggle(ByVal Value As Boolean)
PropDblClickToggle = Value
If CoolBarHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(CoolBarHandle, GWL_STYLE)
    If PropDblClickToggle = True Then
        If Not (dwStyle And RBS_DBLCLKTOGGLE) = RBS_DBLCLKTOGGLE Then dwStyle = dwStyle Or RBS_DBLCLKTOGGLE
    Else
        If (dwStyle And RBS_DBLCLKTOGGLE) = RBS_DBLCLKTOGGLE Then dwStyle = dwStyle And Not RBS_DBLCLKTOGGLE
    End If
    SetWindowLong CoolBarHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "DblClickToggle"
End Property

Public Property Get VerticalGripper() As Boolean
Attribute VerticalGripper.VB_Description = "Returns/sets a value that determines whether or not the size grip will be displayed vertically instead of horizontally in a vertical cool bar."
VerticalGripper = PropVerticalGripper
End Property

Public Property Let VerticalGripper(ByVal Value As Boolean)
PropVerticalGripper = Value
If CoolBarHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(CoolBarHandle, GWL_STYLE)
    If PropVerticalGripper = True Then
        If Not (dwStyle And RBS_VERTICALGRIPPER) = RBS_VERTICALGRIPPER Then dwStyle = dwStyle Or RBS_VERTICALGRIPPER
    Else
        If (dwStyle And RBS_VERTICALGRIPPER) = RBS_VERTICALGRIPPER Then dwStyle = dwStyle And Not RBS_VERTICALGRIPPER
    End If
    SetWindowLong CoolBarHandle, GWL_STYLE, dwStyle
    Call ResetHeaderSizes
    Me.Refresh
End If
UserControl.PropertyChanged "VerticalGripper"
End Property

Public Property Get ShowTips() As Boolean
Attribute ShowTips.VB_Description = "Returns/sets a value that determines whether the tool tip text properties will be displayed or not."
ShowTips = PropShowTips
End Property

Public Property Let ShowTips(ByVal Value As Boolean)
PropShowTips = Value
If CoolBarHandle <> 0 And CoolBarDesignMode = False Then
    If PropShowTips = False Then
        Call DestroyToolTip
    Else
        Call CreateToolTip
    End If
End If
UserControl.PropertyChanged "ShowTips"
End Property

Public Property Get DoubleBuffer() As Boolean
Attribute DoubleBuffer.VB_Description = "Returns/sets a value that determines whether the control paints via double-buffering, which reduces flicker."
DoubleBuffer = PropDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal Value As Boolean)
PropDoubleBuffer = Value
UserControl.PropertyChanged "DoubleBuffer"
End Property

Public Property Get Bands() As CbrBands
Attribute Bands.VB_Description = "Returns a reference to a collection of the band objects."
If PropBands Is Nothing Then
    Set PropBands = New CbrBands
    PropBands.FInit Me, Ambient.UserMode
End If
Set Bands = PropBands
End Property

Friend Sub FBandsAdd(ByVal Index As Long, ByVal NewBand As CbrBand, Optional ByVal Caption As String, Optional ByVal ImageIndex As Long, Optional ByVal NewRow As Boolean, Optional ByVal Child As Object, Optional ByVal Visible As Boolean = True)
Dim RBBI As REBARBANDINFO
With RBBI
.cbSize = LenB(RBBI)
.fMask = RBBIM_ID Or RBBIM_LPARAM Or RBBIM_TEXT Or RBBIM_IMAGE Or RBBIM_STYLE Or RBBIM_CHILD Or RBBIM_CHILDSIZE
.wID = NextBandID()
NewBand.ID = .wID
.iImage = ImageIndex - 1
.lpText = StrPtr(Caption)
.cch = Len(Caption)
.lParam = ObjPtr(NewBand)
.fStyle = RBBS_FIXEDBMP
If NewRow = True Then .fStyle = .fStyle Or RBBS_BREAK
If Visible = False Then .fStyle = .fStyle Or RBBS_HIDDEN
If Not Child Is Nothing Then
    Dim Handle As Long
    On Error Resume Next
    Handle = Child.hWndUserControl
    If Err.Number <> 0 Then Handle = Child.hWnd
    On Error GoTo 0
    Call EvaluateWndChild(Handle)
    If ControlIsValid(Child) = True And Handle <> 0 Then
        .hWndChild = Handle
    Else
        Err.Raise 380
    End If
Else
    .hWndChild = -1
End If
.CXMinChild = 0
.CYMinChild = (24 * PixelsPerDIP_Y())
If Not PropPicture Is Nothing Then
    .fMask = .fMask Or RBBIM_BACKGROUND
    .hBmpBack = PropPicture.Handle
End If
End With
If CoolBarHandle <> 0 Then
    If Index = 0 Then
        SendMessage CoolBarHandle, RB_INSERTBAND, -1, ByVal VarPtr(RBBI)
    Else
        SendMessage CoolBarHandle, RB_INSERTBAND, Index - 1, ByVal VarPtr(RBBI)
    End If
    Me.Refresh
End If
UserControl.PropertyChanged "InitBands"
End Sub

Friend Sub FBandsRemove(ByVal ID As Long)
If CoolBarHandle <> 0 Then
    SendMessage CoolBarHandle, RB_DELETEBAND, SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&), ByVal 0&
    Me.Refresh
End If
UserControl.PropertyChanged "InitBands"
End Sub

Friend Sub FBandsClear()
If CoolBarHandle <> 0 Then
    Do While SendMessage(CoolBarHandle, RB_DELETEBAND, 0, ByVal 0&) <> 0: Loop
    Me.Refresh
End If
End Sub

Friend Function FBandsPositionToIndex(ByVal Position As Long) As Long
If CoolBarHandle <> 0 Then
    Dim RBBI As REBARBANDINFO, Band As CbrBand
    With RBBI
    .cbSize = LenB(RBBI)
    .fMask = RBBIM_LPARAM
    If SendMessage(CoolBarHandle, RB_GETBANDINFO, Position - 1, ByVal VarPtr(RBBI)) <> 0 Then
        If .lParam <> 0 Then
            Set Band = PtrToObj(.lParam)
            FBandsPositionToIndex = Band.Index
        End If
    End If
    End With
End If
End Function

Friend Property Get FBandCaption(ByVal ID As Long) As String
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO, Buffer As String
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_TEXT
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        Buffer = String(.cch, vbNullChar)
        .lpText = StrPtr(Buffer)
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        FBandCaption = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
        End With
    End If
End If
End Property

Friend Property Let FBandCaption(ByVal ID As Long, ByVal Value As String)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_TEXT
        .lpText = StrPtr(Value)
        .cch = Len(Value)
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Let FBandChild(ByVal ID As Long, ByVal Value As Object)
Set Me.FBandChild(ID) = Value
End Property

Friend Property Set FBandChild(ByVal ID As Long, ByVal Value As Object)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim PrevWndChild As Long
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_CHILD
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        PrevWndChild = .hWndChild
        If Not Value Is Nothing Then
            Dim Handle As Long
            On Error Resume Next
            Handle = Value.hWndUserControl
            If Err.Number <> 0 Then Handle = Value.hWnd
            On Error GoTo 0
            Call EvaluateWndChild(Handle)
            If ControlIsValid(Value) = True And Handle <> 0 Then
                .hWndChild = Handle
            Else
                Err.Raise 380
            End If
        Else
            .hWndChild = -1
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        If PrevWndChild <> 0 And PrevWndChild <> .hWndChild And Not PrevWndChild = -1 Then
            SetParent PrevWndChild, UserControl.hWnd
            ShowWindow PrevWndChild, SW_SHOW
        End If
        End With
        Call UserControl_Resize
    End If
End If
End Property

Friend Property Get FBandStyle(ByVal ID As Long) As CbrBandStyleConstants
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (.fStyle And RBBS_FIXEDSIZE) = RBBS_FIXEDSIZE Then
            FBandStyle = CbrBandStyleFixedSize
        Else
            FBandStyle = CbrBandStyleNormal
        End If
        End With
    End If
End If
End Property

Friend Property Let FBandStyle(ByVal ID As Long, ByVal Value As CbrBandStyleConstants)
If CoolBarHandle <> 0 Then
    Select Case Value
        Case CbrBandStyleNormal, CbrBandStyleFixedSize
            Dim Index As Long
            Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
            If Index > -1 Then
                Dim RBBI As REBARBANDINFO
                With RBBI
                .cbSize = LenB(RBBI)
                .fMask = RBBIM_STYLE
                SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
                If Value = CbrBandStyleFixedSize Then
                    If Not (.fStyle And RBBS_FIXEDSIZE) = RBBS_FIXEDSIZE Then .fStyle = .fStyle Or RBBS_FIXEDSIZE
                ElseIf Value = CbrBandStyleNormal Then
                    If (.fStyle And RBBS_FIXEDSIZE) = RBBS_FIXEDSIZE Then .fStyle = .fStyle And Not RBBS_FIXEDSIZE
                End If
                SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
                End With
            End If
        Case Else
            Err.Raise 380
    End Select
End If
End Property

Friend Property Get FBandImage(ByVal ID As Long) As Long
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_IMAGE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        FBandImage = .iImage + 1
        End With
    End If
End If
End Property

Friend Property Let FBandImage(ByVal ID As Long, ByVal Value As Long)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_IMAGE
        .iImage = Value - 1
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandWidth(ByVal ID As Long) As Single
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_SIZE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
            FBandWidth = UserControl.ScaleX(.CX, vbPixels, vbContainerSize)
        Else
            FBandWidth = UserControl.ScaleY(.CX, vbPixels, vbContainerSize)
        End If
        End With
    End If
End If
End Property

Friend Property Let FBandWidth(ByVal ID As Long, ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If CoolBarHandle <> 0 Then
    If Me.FBandStyle(ID) <> CbrBandStyleFixedSize Then
        Dim Index As Long
        Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
        If Index > -1 Then
            Dim RBBI As REBARBANDINFO
            With RBBI
            .cbSize = LenB(RBBI)
            .fMask = RBBIM_SIZE
            SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
            If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
                .CX = UserControl.ScaleX(Value, vbContainerSize, vbPixels)
            Else
                .CX = UserControl.ScaleY(Value, vbContainerSize, vbPixels)
            End If
            SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
            End With
        End If
    Else
        Err.Raise Number:=35800, Description:="Property is read-only if the style property is set to fixed size"
    End If
End If
End Property

Friend Property Get FBandHeight(ByVal ID As Long) As Single
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RC As RECT
        SendMessage CoolBarHandle, RB_GETRECT, Index, ByVal VarPtr(RC)
        If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
            FBandHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
        Else
            FBandHeight = UserControl.ScaleX((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
        End If
    End If
End If
End Property

Friend Property Get FBandMinWidth(ByVal ID As Long) As Single
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_CHILDSIZE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
            FBandMinWidth = UserControl.ScaleX(.CXMinChild, vbPixels, vbContainerSize)
        Else
            FBandMinWidth = UserControl.ScaleY(.CXMinChild, vbPixels, vbContainerSize)
        End If
        End With
    End If
End If
End Property

Friend Property Let FBandMinWidth(ByVal ID As Long, ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE Or RBBIM_CHILDSIZE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
            .CXMinChild = UserControl.ScaleX(Value, vbContainerSize, vbPixels)
        Else
            .CXMinChild = UserControl.ScaleY(Value, vbContainerSize, vbPixels)
        End If
        If (.fStyle And RBBS_FIXEDSIZE) = RBBS_FIXEDSIZE Then .CXMinChild = .CXMinChild + 8
        .fMask = RBBIM_CHILDSIZE
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandMinHeight(ByVal ID As Long) As Single
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_CHILDSIZE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
            FBandMinHeight = UserControl.ScaleY(.CYMinChild, vbPixels, vbContainerSize)
        Else
            FBandMinHeight = UserControl.ScaleX(.CYMinChild, vbPixels, vbContainerSize)
        End If
        End With
    End If
End If
End Property

Friend Property Let FBandMinHeight(ByVal ID As Long, ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_CHILDSIZE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
            .CYMinChild = UserControl.ScaleY(Value, vbContainerSize, vbPixels)
        Else
            .CYMinChild = UserControl.ScaleX(Value, vbContainerSize, vbPixels)
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandIdealWidth(ByVal ID As Long) As Single
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_IDEALSIZE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
            FBandIdealWidth = UserControl.ScaleX(.CXIdeal, vbPixels, vbContainerSize)
        Else
            FBandIdealWidth = UserControl.ScaleY(.CXIdeal, vbPixels, vbContainerSize)
        End If
        End With
    End If
End If
End Property

Friend Property Let FBandIdealWidth(ByVal ID As Long, ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_IDEALSIZE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
            .CXIdeal = UserControl.ScaleX(Value, vbContainerSize, vbPixels)
        Else
            .CXIdeal = UserControl.ScaleY(Value, vbContainerSize, vbPixels)
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandGripper(ByVal ID As Long) As CbrBandGripperConstants
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If (.fStyle And RBBS_NOGRIPPER) = RBBS_NOGRIPPER Then
            FBandGripper = CbrBandGripperNever
        ElseIf (.fStyle And RBBS_GRIPPERALWAYS) = RBBS_GRIPPERALWAYS Then
            FBandGripper = CbrBandGripperAlways
        Else
            FBandGripper = CbrBandGripperNormal
        End If
        End With
    End If
End If
End Property

Friend Property Let FBandGripper(ByVal ID As Long, ByVal Value As CbrBandGripperConstants)
If CoolBarHandle <> 0 Then
    Select Case Value
        Case CbrBandGripperNormal, CbrBandGripperAlways, CbrBandGripperNever
            Dim Index As Long
            Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
            If Index > -1 Then
                Dim RBBI As REBARBANDINFO
                With RBBI
                .cbSize = LenB(RBBI)
                .fMask = RBBIM_STYLE
                SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
                If Value = CbrBandGripperNever Then
                    If Not (.fStyle And RBBS_NOGRIPPER) = RBBS_NOGRIPPER Then .fStyle = .fStyle Or RBBS_NOGRIPPER
                    If (.fStyle And RBBS_GRIPPERALWAYS) = RBBS_GRIPPERALWAYS Then .fStyle = .fStyle And Not RBBS_GRIPPERALWAYS
                ElseIf Value = CbrBandGripperAlways Then
                    If (.fStyle And RBBS_NOGRIPPER) = RBBS_NOGRIPPER Then .fStyle = .fStyle And Not RBBS_NOGRIPPER
                    If Not (.fStyle And RBBS_GRIPPERALWAYS) = RBBS_GRIPPERALWAYS Then .fStyle = .fStyle Or RBBS_GRIPPERALWAYS
                ElseIf Value = CbrBandGripperNormal Then
                    If (.fStyle And RBBS_NOGRIPPER) = RBBS_NOGRIPPER Then .fStyle = .fStyle And Not RBBS_NOGRIPPER
                    If (.fStyle And RBBS_GRIPPERALWAYS) = RBBS_GRIPPERALWAYS Then .fStyle = .fStyle And Not RBBS_GRIPPERALWAYS
                End If
                SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
                End With
            End If
        Case Else
            Err.Raise 380
    End Select
End If
End Property

Friend Property Let FBandUseCoolBarPicture(ByVal ID As Long, ByVal Picture As IPictureDisp, ByVal Value As Boolean)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_BACKGROUND
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value = False Then
            If Picture Is Nothing Then
                .hBmpBack = 0
            Else
                .hBmpBack = Picture.Handle
            End If
        Else
            If PropPicture Is Nothing Then
                .hBmpBack = 0
            Else
                .hBmpBack = PropPicture.Handle
            End If
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Let FBandPicture(ByVal ID As Long, ByVal Value As IPictureDisp)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_BACKGROUND
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value Is Nothing Then
            .hBmpBack = 0
        Else
            .hBmpBack = Value.Handle
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Let FBandUseCoolBarColors(ByVal ID As Long, ByVal BackColor As OLE_COLOR, ByVal ForeColor As OLE_COLOR, ByVal Value As Boolean)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_COLORS
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value = False Then
            .BackColor = WinColor(BackColor)
            .ForeColor = WinColor(ForeColor)
        Else
            .BackColor = -1
            .ForeColor = -1
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Let FBandBackColor(ByVal ID As Long, ByVal Value As OLE_COLOR)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_COLORS
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        .BackColor = WinColor(Value)
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Let FBandForeColor(ByVal ID As Long, ByVal Value As OLE_COLOR)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_COLORS
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        .ForeColor = WinColor(Value)
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandNewRow(ByVal ID As Long) As Boolean
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        FBandNewRow = CBool((.fStyle And RBBS_BREAK) = RBBS_BREAK)
        End With
    End If
End If
End Property

Friend Property Let FBandNewRow(ByVal ID As Long, ByVal Value As Boolean)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value = True Then
            If Not (.fStyle And RBBS_BREAK) = RBBS_BREAK Then .fStyle = .fStyle Or RBBS_BREAK
        Else
            If (.fStyle And RBBS_BREAK) = RBBS_BREAK Then .fStyle = .fStyle And Not RBBS_BREAK
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandVisible(ByVal ID As Long) As Boolean
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        FBandVisible = CBool((.fStyle And RBBS_HIDDEN) = 0)
        End With
    End If
End If
End Property

Friend Property Let FBandVisible(ByVal ID As Long, ByVal Value As Boolean)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value = True Then
            If (.fStyle And RBBS_HIDDEN) = RBBS_HIDDEN Then .fStyle = .fStyle And Not RBBS_HIDDEN
        Else
            If Not (.fStyle And RBBS_HIDDEN) = RBBS_HIDDEN Then .fStyle = .fStyle Or RBBS_HIDDEN
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandChildEdge(ByVal ID As Long) As Boolean
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        FBandChildEdge = CBool((.fStyle And RBBS_CHILDEDGE) <> 0)
        End With
    End If
End If
End Property

Friend Property Let FBandChildEdge(ByVal ID As Long, ByVal Value As Boolean)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value = True Then
            If Not (.fStyle And RBBS_CHILDEDGE) = RBBS_CHILDEDGE Then .fStyle = .fStyle Or RBBS_CHILDEDGE
        Else
            If (.fStyle And RBBS_CHILDEDGE) = RBBS_CHILDEDGE Then .fStyle = .fStyle Or RBBS_CHILDEDGE
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandUseChevron(ByVal ID As Long) As Boolean
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        FBandUseChevron = CBool((.fStyle And RBBS_USECHEVRON) = RBBS_USECHEVRON)
        End With
    End If
End If
End Property

Friend Property Let FBandUseChevron(ByVal ID As Long, ByVal Value As Boolean)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE Or RBBIM_HEADERSIZE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value = True Then
            If Not (.fStyle And RBBS_USECHEVRON) = RBBS_USECHEVRON Then .fStyle = .fStyle Or RBBS_USECHEVRON
        Else
            If (.fStyle And RBBS_USECHEVRON) = RBBS_USECHEVRON Then .fStyle = .fStyle And Not RBBS_USECHEVRON
        End If
        .CXHeader = -1
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandHideCaption(ByVal ID As Long) As Boolean
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        FBandHideCaption = CBool((.fStyle And RBBS_HIDETITLE) = RBBS_HIDETITLE)
        End With
    End If
End If
End Property

Friend Property Let FBandHideCaption(ByVal ID As Long, ByVal Value As Boolean)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value = True Then
            If Not (.fStyle And RBBS_HIDETITLE) = RBBS_HIDETITLE Then .fStyle = .fStyle Or RBBS_HIDETITLE
        Else
            If (.fStyle And RBBS_HIDETITLE) = RBBS_HIDETITLE Then .fStyle = .fStyle And Not RBBS_HIDETITLE
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandFixedBackground(ByVal ID As Long) As Boolean
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        FBandFixedBackground = CBool((.fStyle And RBBS_FIXEDBMP) = RBBS_FIXEDBMP)
        End With
    End If
End If
End Property

Friend Property Let FBandFixedBackground(ByVal ID As Long, ByVal Value As Boolean)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_STYLE
        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
        If Value = True Then
            If Not (.fStyle And RBBS_FIXEDBMP) = RBBS_FIXEDBMP Then .fStyle = .fStyle Or RBBS_FIXEDBMP
        Else
            If (.fStyle And RBBS_FIXEDBMP) = RBBS_FIXEDBMP Then .fStyle = .fStyle And Not RBBS_FIXEDBMP
        End If
        SendMessage CoolBarHandle, RB_SETBANDINFO, Index, ByVal VarPtr(RBBI)
        End With
    End If
End If
End Property

Friend Property Get FBandPosition(ByVal ID As Long) As Long
If CoolBarHandle <> 0 Then FBandPosition = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&) + 1
End Property

Friend Property Let FBandPosition(ByVal ID As Long, ByVal Value As Long)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then If SendMessage(CoolBarHandle, RB_MOVEBAND, Index, ByVal CLng(Value - 1)) = 0 Then Err.Raise 380
End If
End Property

Friend Sub FBandMaximize(ByVal ID As Long)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then SendMessage CoolBarHandle, RB_MAXIMIZEBAND, Index, ByVal 1&
End If
End Sub

Friend Sub FBandMinimize(ByVal ID As Long)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then SendMessage CoolBarHandle, RB_MINIMIZEBAND, Index, ByVal 0&
End If
End Sub

Friend Sub FBandPushChevron(ByVal ID As Long)
If CoolBarHandle <> 0 Then
    Dim Index As Long
    Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, ID, ByVal 0&)
    If Index > -1 Then SendMessage CoolBarHandle, RB_PUSHCHEVRON, Index, ByVal 0&
End If
End Sub

Private Sub CreateCoolBar()
If CoolBarHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or CCS_NODIVIDER Or CCS_NORESIZE
dwExStyle = WS_EX_TOOLWINDOW
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If PropBorderStyle <> vbBSNone Then dwStyle = dwStyle Or WS_BORDER
If PropOrientation = CbrOrientationVertical Then dwStyle = dwStyle Or CCS_VERT
If PropBandBorders = True Then dwStyle = dwStyle Or RBS_BANDBORDERS
If PropFixedOrder = True Then dwStyle = dwStyle Or RBS_FIXEDORDER
If PropVariantHeight = True Then dwStyle = dwStyle Or RBS_VARHEIGHT
If PropDblClickToggle = True Then dwStyle = dwStyle Or RBS_DBLCLKTOGGLE
If PropVerticalGripper = True Then dwStyle = dwStyle Or RBS_VERTICALGRIPPER
If CoolBarDesignMode = False Then
    ' The WM_NOTIFYFORMAT notification must be handled, which will be sent on control creation.
    ' Thus it is necessary to subclass the parent before the control is created.
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
CoolBarHandle = CreateWindowEx(dwExStyle, StrPtr("ReBarWindow32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.BackColor = PropBackColor
Me.ForeColor = PropForeColor
Me.ShowTips = PropShowTips
If CoolBarHandle <> 0 Then
    If ComCtlsSupportLevel() = 0 Then
        ' The '&' won't underline the character, instead it is displayed as a normal character.
        ' This behavior is necessary for backward compatibility with earlier versions of the common controls.
        ' If you want the character to be underlined, it is necessary to send a CCM_SETVERSION message
        ' with the wParam value set to 5 before adding any items to the control.
        SendMessage CoolBarHandle, CCM_SETVERSION, 5, ByVal 0&
    End If
End If
If CoolBarDesignMode = False Then
    If CoolBarHandle <> 0 Then Call ComCtlsSetSubclass(CoolBarHandle, Me, 1)
Else
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
End If
End Sub

Private Sub CreateToolTip()
Static Done As Boolean
Dim dwExStyle As Long
If CoolBarToolTipHandle <> 0 Then Exit Sub
If Done = False Then
    Call ComCtlsInitCC(ICC_TAB_CLASSES)
    Done = True
End If
dwExStyle = WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
CoolBarToolTipHandle = CreateWindowEx(dwExStyle, StrPtr("tooltips_class32"), StrPtr("Tool Tip"), WS_POPUP Or TTS_ALWAYSTIP Or TTS_NOPREFIX, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If CoolBarToolTipHandle <> 0 Then
    Call ComCtlsInitToolTip(CoolBarToolTipHandle)
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = CoolBarHandle
    .uId = 0
    .uFlags = TTF_SUBCLASS Or TTF_PARSELINKS
    If PropRightToLeft = True And PropRightToLeftLayout = False Then .uFlags = .uFlags Or TTF_RTLREADING
    .lpszText = LPSTR_TEXTCALLBACK
    GetClientRect CoolBarHandle, .RC
    End With
    SendMessage CoolBarToolTipHandle, TTM_ADDTOOL, 0, ByVal VarPtr(TI)
End If
Call SetVisualStylesToolTip
End Sub

Private Sub DestroyCoolBar()
If CoolBarHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(CoolBarHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
Call DestroyToolTip
ShowWindow CoolBarHandle, SW_HIDE
SetParent CoolBarHandle, 0
DestroyWindow CoolBarHandle
CoolBarHandle = 0
If CoolBarFontHandle <> 0 Then
    DeleteObject CoolBarFontHandle
    CoolBarFontHandle = 0
End If

#If ImplementThemedReBarFix = True Then

If CoolBarTheme <> 0 Then
    CloseThemeData CoolBarTheme
    CoolBarTheme = 0
End If

#End If

End Sub

Private Sub DestroyToolTip()
If CoolBarToolTipHandle = 0 Then Exit Sub
DestroyWindow CoolBarToolTipHandle
CoolBarToolTipHandle = 0
CoolBarToolTipIndex = -1
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get ContainedControls() As VBRUN.ContainedControls
Attribute ContainedControls.VB_Description = "Returns a collection that allows access to the controls contained within the control that were added to the control by the developer who uses the control."
Set ContainedControls = UserControl.ContainedControls
End Property

Public Property Get RowCount() As Long
Attribute RowCount.VB_Description = "Returns the number of rows."
Attribute RowCount.VB_MemberFlags = "400"
If CoolBarHandle <> 0 Then RowCount = SendMessage(CoolBarHandle, RB_GETROWCOUNT, 0, ByVal 0&)
End Property

Public Function HitTest(ByVal X As Single, ByVal Y As Single, Optional ByRef HitResult As CbrHitResultConstants) As CbrBand
Attribute HitTest.VB_Description = "Returns a reference to the band object located at the coordinates of X and Y."
If CoolBarHandle <> 0 Then
    Dim RBHTI As RBHITTESTINFO
    With RBHTI
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    If SendMessage(CoolBarHandle, RB_HITTEST, 0, ByVal VarPtr(RBHTI)) > -1 Then
        Select Case .Flag
            Case RBHT_NOWHERE
                HitResult = CbrHitResultNoWhere
            Case RBHT_CAPTION
                HitResult = CbrHitResultCaption
            Case RBHT_CLIENT
                HitResult = CbrHitResultClient
            Case RBHT_GRABBER
                HitResult = CbrHitResultGrabber
            Case RBHT_CHEVRON
                HitResult = CbrHitResultChevron
            Case RBHT_SPLITTER
                HitResult = CbrHitResultSplitter
        End Select
        If .uBand > -1 Then
            Dim RBBI As REBARBANDINFO
            RBBI.cbSize = LenB(RBBI)
            RBBI.fMask = RBBIM_LPARAM
            SendMessage CoolBarHandle, RB_GETBANDINFO, .uBand, ByVal VarPtr(RBBI)
            If RBBI.lParam <> 0 Then Set HitTest = PtrToObj(RBBI.lParam)
        End If
    End If
    End With
End If
End Function

Private Function NextBandID() As Long
Static ID As Long
ID = ID + 1
NextBandID = ID
End Function

Private Sub EvaluateWndChild(ByVal Handle As Long)
If CoolBarHandle <> 0 And Handle <> 0 Then
    Dim Count As Long
    Count = SendMessage(CoolBarHandle, RB_GETBANDCOUNT, 0, ByVal 0&)
    If Count > 0 Then
        Dim i As Long
        Dim RBBI As REBARBANDINFO, Band As CbrBand
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_CHILD Or RBBIM_LPARAM
        For i = 0 To Count - 1
            SendMessage CoolBarHandle, RB_GETBANDINFO, i, ByVal VarPtr(RBBI)
            If .hWndChild = Handle And Not .hWndChild = -1 Then
                If .lParam <> 0 Then
                    Set Band = PtrToObj(.lParam)
                    Band.FInit Me, Band.Key, Nothing, Band.Image, Band.ImageIndex
                End If
                .fMask = RBBIM_CHILD
                .hWndChild = -1
                SendMessage CoolBarHandle, RB_SETBANDINFO, i, ByVal VarPtr(RBBI)
                SetParent Handle, UserControl.hWnd
            End If
        Next i
        End With
    End If
End If
End Sub

Private Sub ResetHeaderSizes()
If CoolBarHandle <> 0 Then
    Dim Count As Long
    Count = SendMessage(CoolBarHandle, RB_GETBANDCOUNT, 0, ByVal 0&)
    If Count > 0 Then
        Dim i As Long
        Dim RBBI As REBARBANDINFO
        With RBBI
        .cbSize = LenB(RBBI)
        .fMask = RBBIM_HEADERSIZE
        For i = 0 To Count - 1
            SendMessage CoolBarHandle, RB_GETBANDINFO, i, ByVal VarPtr(RBBI)
            .CXHeader = -1
            SendMessage CoolBarHandle, RB_SETBANDINFO, i, ByVal VarPtr(RBBI)
        Next i
        End With
    End If
End If
End Sub

Private Sub SetVisualStylesToolTip()
If CoolBarHandle <> 0 Then
    If CoolBarToolTipHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles CoolBarToolTipHandle
        Else
            RemoveVisualStyles CoolBarToolTipHandle
        End If
    End If
End If
End Sub

Private Sub UpdateToolTipRect()
If CoolBarHandle <> 0 And CoolBarToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = CoolBarHandle
    .uId = 0
    GetClientRect CoolBarHandle, .RC
    SendMessage CoolBarToolTipHandle, TTM_NEWTOOLRECT, 0, ByVal VarPtr(TI)
    End With
End If
End Sub

Private Sub CheckToolTipIndex(ByVal X As Long, ByVal Y As Long)
If CoolBarHandle <> 0 And CoolBarToolTipHandle <> 0 Then
    Dim RBHTI As RBHITTESTINFO
    With RBHTI
    .PT.X = X
    .PT.Y = Y
    If SendMessage(CoolBarHandle, RB_HITTEST, 0, ByVal VarPtr(RBHTI)) > -1 Then
        If .Flag = RBHT_CAPTION And .uBand > -1 Then
            If CoolBarToolTipIndex <> .uBand Then
                CoolBarToolTipIndex = .uBand
                SendMessage CoolBarToolTipHandle, TTM_POP, 0, ByVal 0&
            End If
        Else
            CoolBarToolTipIndex = -1
            SendMessage CoolBarToolTipHandle, TTM_POP, 0, ByVal 0&
        End If
    Else
        CoolBarToolTipIndex = -1
        SendMessage CoolBarToolTipHandle, TTM_POP, 0, ByVal 0&
    End If
    End With
End If
End Sub

Private Function ControlIsValid(ByVal Control As Object) As Boolean
On Error Resume Next
Dim Container As Object
Set Container = Control.Container
ControlIsValid = CBool(Err.Number = 0 And Not Control Is Extender And Container Is Extender)
On Error GoTo 0
End Function

Private Function PropImageListControl() As Object
If CoolBarImageListObjectPointer <> 0 Then Set PropImageListControl = PtrToObj(CoolBarImageListObjectPointer)
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
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            Dim hCursor As Long
            If MousePointerID(PropMousePointer) <> 0 Then
                hCursor = LoadCursor(0, MousePointerID(PropMousePointer))
            ElseIf PropMousePointer = 99 Then
                If Not PropMouseIcon Is Nothing Then hCursor = PropMouseIcon.Handle
            End If
            If hCursor <> 0 Then
                SetCursor hCursor
                WindowProcControl = 1
                Exit Function
            ElseIf hWnd <> wParam And wParam <> 0 Then
                ' Ensures that the cild controls can walk up the chain properly.
                WindowProcControl = DefWindowProc(hWnd, wMsg, wParam, lParam)
                Exit Function
            End If
        End If
    Case WM_ERASEBKGND
        If PropDoubleBuffer = True And (CoolBarDoubleBufferEraseBkgDC <> wParam Or CoolBarDoubleBufferEraseBkgDC = 0) And WindowFromDC(wParam) = hWnd Then
            WindowProcControl = 0
            Exit Function
        End If
    Case WM_PAINT
        If PropDoubleBuffer = True Then
            Dim ClientRect As RECT, hDC As Long
            Dim hDCBmp As Long
            Dim hBmp As Long, hBmpOld As Long
            GetClientRect hWnd, ClientRect
            Dim PS As PAINTSTRUCT
            hDC = BeginPaint(hWnd, PS)
            With PS
            If wParam <> 0 Then hDC = wParam
            hDCBmp = CreateCompatibleDC(hDC)
            If hDCBmp <> 0 Then
                hBmp = CreateCompatibleBitmap(hDC, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top)
                If hBmp <> 0 Then
                    hBmpOld = SelectObject(hDCBmp, hBmp)
                    CoolBarDoubleBufferEraseBkgDC = hDCBmp
                    SendMessage hWnd, WM_PRINT, hDCBmp, ByVal PRF_CLIENT Or PRF_ERASEBKGND
                    CoolBarDoubleBufferEraseBkgDC = 0
                    With PS.RCPaint
                    BitBlt hDC, .Left, .Top, .Right - .Left, .Bottom - .Top, hDCBmp, .Left, .Top, vbSrcCopy
                    End With
                    SelectObject hDCBmp, hBmpOld
                    DeleteObject hBmp
                End If
                DeleteDC hDCBmp
            End If
            End With
            EndPaint hWnd, PS
            WindowProcControl = 0
            Exit Function
        End If
    Case WM_SIZE
        If PropShowTips = True Then Call UpdateToolTipRect
    Case WM_MOUSEMOVE
        If PropShowTips = True Then Call CheckToolTipIndex(Get_X_lParam(lParam), Get_Y_lParam(lParam))
    Case WM_NOTIFY
        Dim NM As NMHDR
        Dim RBBI As REBARBANDINFO, Band As CbrBand
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = CoolBarToolTipHandle And CoolBarToolTipHandle <> 0 Then
            Select Case NM.Code
                Case TTN_GETDISPINFO
                    Dim NMTTDI As NMTTDISPINFO
                    CopyMemory NMTTDI, ByVal lParam, LenB(NMTTDI)
                    With NMTTDI
                    Dim Text As String, RBHTI As RBHITTESTINFO, Pos As Long
                    With RBHTI
                    Pos = GetMessagePos()
                    .PT.X = Get_X_lParam(Pos)
                    .PT.Y = Get_Y_lParam(Pos)
                    ScreenToClient hWnd, .PT
                    If SendMessage(hWnd, RB_HITTEST, 0, ByVal VarPtr(RBHTI)) > -1 Then
                        If .Flag = RBHT_CAPTION And .uBand > -1 Then
                            RBBI.cbSize = LenB(RBBI)
                            RBBI.fMask = RBBIM_LPARAM
                            SendMessage hWnd, RB_GETBANDINFO, .uBand, ByVal VarPtr(RBBI)
                            If RBBI.lParam <> 0 Then
                                Set Band = PtrToObj(RBBI.lParam)
                                Text = Band.ToolTipText
                            End If
                        End If
                    End If
                    End With
                    If Not Text = vbNullString Then
                        If Len(Text) <= 80 Then
                            Text = Left$(Text & vbNullChar, 80)
                            CopyMemory .szText(0), ByVal StrPtr(Text), LenB(Text)
                        Else
                            .lpszText = StrPtr(Text)
                        End If
                        .hInst = 0
                        CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                    End If
                    End With
            End Select
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
                CoolBarIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                CoolBarIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                CoolBarIsClick = True
            Case WM_MOUSEMOVE
                If CoolBarMouseOver = False And PropMouseTrack = True Then
                    CoolBarMouseOver = True
                    RaiseEvent MouseEnter
                    Dim RBHTI1 As RBHITTESTINFO
                    With RBHTI1
                    .PT.X = Get_X_lParam(lParam)
                    .PT.Y = Get_Y_lParam(lParam)
                    CoolBarMouseOverIndex = SendMessage(CoolBarHandle, RB_HITTEST, 0, ByVal VarPtr(RBHTI1)) + 1
                    If CoolBarMouseOverIndex > 0 Then
                        If .uBand > -1 Then
                            Dim RBBI1 As REBARBANDINFO
                            RBBI1.cbSize = LenB(RBBI1)
                            RBBI1.fMask = RBBIM_LPARAM
                            SendMessage CoolBarHandle, RB_GETBANDINFO, .uBand, ByVal VarPtr(RBBI1)
                            If RBBI1.lParam <> 0 Then RaiseEvent BandMouseEnter(PtrToObj(RBBI1.lParam))
                        End If
                    End If
                    End With
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                If CoolBarMouseOver = True And PropMouseTrack = True Then
                    Dim RBHTI2 As RBHITTESTINFO, Index As Long
                    With RBHTI2
                    .PT.X = Get_X_lParam(lParam)
                    .PT.Y = Get_Y_lParam(lParam)
                    Index = SendMessage(CoolBarHandle, RB_HITTEST, 0, ByVal VarPtr(RBHTI2)) + 1
                    If CoolBarMouseOverIndex <> Index Then
                        If CoolBarMouseOverIndex > 0 Then
                            Dim RBBI2 As REBARBANDINFO
                            RBBI2.cbSize = LenB(RBBI2)
                            RBBI2.fMask = RBBIM_LPARAM
                            SendMessage CoolBarHandle, RB_GETBANDINFO, CoolBarMouseOverIndex - 1, ByVal VarPtr(RBBI2)
                            If RBBI2.lParam <> 0 Then RaiseEvent BandMouseLeave(PtrToObj(RBBI2.lParam))
                        End If
                        CoolBarMouseOverIndex = Index
                        If CoolBarMouseOverIndex > 0 Then
                            Dim RBBI3 As REBARBANDINFO
                            RBBI3.cbSize = LenB(RBBI3)
                            RBBI3.fMask = RBBIM_LPARAM
                            SendMessage CoolBarHandle, RB_GETBANDINFO, CoolBarMouseOverIndex - 1, ByVal VarPtr(RBBI3)
                            If RBBI3.lParam <> 0 Then RaiseEvent BandMouseEnter(PtrToObj(RBBI3.lParam))
                        End If
                    End If
                    End With
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
                If CoolBarIsClick = True Then
                    CoolBarIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If CoolBarMouseOver = True Then
            CoolBarMouseOver = False
            If CoolBarMouseOverIndex > 0 Then
                Dim RBBI4 As REBARBANDINFO
                RBBI4.cbSize = LenB(RBBI4)
                RBBI4.fMask = RBBIM_LPARAM
                SendMessage CoolBarHandle, RB_GETBANDINFO, CoolBarMouseOverIndex - 1, ByVal VarPtr(RBBI4)
                If RBBI4.lParam <> 0 Then RaiseEvent BandMouseLeave(PtrToObj(RBBI4.lParam))
            End If
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR, NMRB As NMREBAR
        Dim RBBI As REBARBANDINFO, Cancel As Boolean, Band As CbrBand
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = CoolBarHandle Then
            Select Case NM.Code
                Case NM_CUSTOMDRAW
                    Dim NMCD As NMCUSTOMDRAW
                    CopyMemory NMCD, ByVal lParam, LenB(NMCD)
                    Select Case NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            If PropRightToLeft = True And PropRightToLeftLayout = False Then
                                Dim fMode As Long
                                fMode = GetTextAlign(NMCD.hDC)
                                If (fMode And TA_RTLREADING) = 0 Then fMode = fMode Or TA_RTLREADING
                                SetTextAlign NMCD.hDC, fMode
                            End If
                            
                            #If ImplementThemedReBarFix = True Then
                            
                            If CoolBarTheme <> 0 Then
                                CloseThemeData CoolBarTheme
                                CoolBarTheme = 0
                            End If
                            If EnabledVisualStyles() = True And PropVisualStyles = True Then
                                If ComCtlsSupportLevel() >= 2 Then CoolBarTheme = OpenThemeData(CoolBarHandle, StrPtr("ReBar"))
                            End If
                            If CoolBarTheme <> 0 Then
                                WindowProcUserControl = CDRF_NOTIFYITEMDRAW Or CDRF_NOTIFYPOSTPAINT
                            Else
                                WindowProcUserControl = CDRF_NOTIFYPOSTPAINT
                            End If
                            
                            #Else
                            
                            WindowProcUserControl = CDRF_DODEFAULT
                            
                            #End If
                            
                            Exit Function
                        
                        #If ImplementThemedReBarFix = True Then
                        
                        Case CDDS_POSTPAINT
                            If CoolBarTheme <> 0 Then
                                CloseThemeData CoolBarTheme
                                CoolBarTheme = 0
                            End If
                        Case CDDS_ITEMPREPAINT
                            If CoolBarTheme <> 0 Then
                                Dim Index As Long, dwStyle As Long
                                With NMCD
                                Index = SendMessage(CoolBarHandle, RB_IDTOINDEX, .dwItemSpec, ByVal 0&)
                                dwStyle = GetWindowLong(CoolBarHandle, GWL_STYLE)
                                Dim RBHTI As RBHITTESTINFO, i As Long, GrabberRect As RECT
                                If (dwStyle And CCS_VERT) = CCS_VERT And Not (dwStyle And RBS_VERTICALGRIPPER) = RBS_VERTICALGRIPPER Then
                                    RBHTI.PT.X = .RC.Left
                                    For i = .RC.Top To .RC.Bottom
                                        RBHTI.PT.Y = i
                                        If SendMessage(CoolBarHandle, RB_HITTEST, 0, ByVal VarPtr(RBHTI)) = Index Then
                                            If RBHTI.Flag <> RBHT_GRABBER Then Exit For
                                        Else
                                            Exit For
                                        End If
                                    Next i
                                    If i > .RC.Top Then
                                        SetRect GrabberRect, .RC.Left, .RC.Top, .RC.Right, i - 1
                                        If IsThemeBackgroundPartiallyTransparent(CoolBarTheme, RP_GRIPPERVERT, 0) <> 0 Then DrawThemeParentBackground CoolBarHandle, .hDC, GrabberRect
                                        DrawThemeBackground CoolBarTheme, .hDC, RP_GRIPPERVERT, 0, GrabberRect, GrabberRect
                                    End If
                                    .RC.Top = i + (1 * PixelsPerDIP_Y())
                                Else
                                    RBHTI.PT.Y = .RC.Top
                                    For i = .RC.Left To .RC.Right
                                        RBHTI.PT.X = i
                                        If SendMessage(CoolBarHandle, RB_HITTEST, 0, ByVal VarPtr(RBHTI)) = Index Then
                                            If RBHTI.Flag <> RBHT_GRABBER Then Exit For
                                        Else
                                            Exit For
                                        End If
                                    Next i
                                    If i > .RC.Left Then
                                        SetRect GrabberRect, .RC.Left, .RC.Top, i - 1, .RC.Bottom
                                        If IsThemeBackgroundPartiallyTransparent(CoolBarTheme, RP_GRIPPER, 0) <> 0 Then DrawThemeParentBackground CoolBarHandle, .hDC, GrabberRect
                                        DrawThemeBackground CoolBarTheme, .hDC, RP_GRIPPER, 0, GrabberRect, GrabberRect
                                    End If
                                    .RC.Left = i + (1 * PixelsPerDIP_X())
                                End If
                                RBBI.cbSize = LenB(RBBI)
                                RBBI.fMask = RBBIM_IMAGE Or RBBIM_STYLE
                                SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI)
                                If RBBI.iImage > -1 Then
                                    Dim RBI As REBARINFO, ImageWidth As Long, ImageHeight As Long
                                    RBI.cbSize = LenB(RBI)
                                    RBI.fMask = RBIM_IMAGELIST
                                    SendMessage CoolBarHandle, RB_GETBARINFO, 0, ByVal VarPtr(RBI)
                                    If RBI.hImageList <> 0 Then
                                        ImageList_GetIconSize RBI.hImageList, ImageWidth, ImageHeight
                                        If (dwStyle And CCS_VERT) = CCS_VERT And Not (dwStyle And RBS_VERTICALGRIPPER) = RBS_VERTICALGRIPPER Then
                                            .RC.Top = .RC.Top + (2 * PixelsPerDIP_Y())
                                            ImageList_Draw RBI.hImageList, RBBI.iImage, .hDC, .RC.Left + ((.RC.Right - .RC.Left - ImageWidth) / 2), .RC.Top, ILD_TRANSPARENT
                                            .RC.Top = .RC.Top + ImageHeight
                                        Else
                                            .RC.Left = .RC.Left + (2 * PixelsPerDIP_X())
                                            ImageList_Draw RBI.hImageList, RBBI.iImage, .hDC, .RC.Left, .RC.Top + ((.RC.Bottom - .RC.Top - ImageHeight) / 2), ILD_TRANSPARENT
                                            .RC.Left = .RC.Left + ImageWidth
                                        End If
                                    End If
                                End If
                                If (RBBI.fStyle And RBBS_HIDETITLE) = 0 Then
                                    Dim Text As String
                                    Text = Me.FBandCaption(.dwItemSpec)
                                    If Not Text = vbNullString Then
                                        Dim hFont As Long, hFontOld As Long, OldBkMode As Long, Format As Long
                                        hFont = SendMessage(CoolBarHandle, WM_GETFONT, 0, ByVal 0&)
                                        If hFont <> 0 Then hFontOld = SelectObject(.hDC, hFont)
                                        OldBkMode = SetBkMode(.hDC, 1)
                                        If (dwStyle And CCS_VERT) = CCS_VERT And Not (dwStyle And RBS_VERTICALGRIPPER) = RBS_VERTICALGRIPPER Then
                                            .RC.Top = .RC.Top + (2 * PixelsPerDIP_Y())
                                            Format = DT_SINGLELINE Or DT_CENTER Or DT_TOP Or DT_END_ELLIPSIS
                                        Else
                                            .RC.Left = .RC.Left + (2 * PixelsPerDIP_X())
                                            Format = DT_SINGLELINE Or DT_LEFT Or DT_VCENTER
                                        End If
                                        If ComCtlsSupportLevel() >= 2 Then
                                            Dim DTTO As DTTOPTS
                                            DTTO.dwSize = LenB(DTTO)
                                            DTTO.dwFlags = DTT_TEXTCOLOR
                                            DTTO.crText = GetTextColor(.hDC)
                                            DrawThemeTextEx CoolBarTheme, .hDC, RP_BAND, 0, StrPtr(Text), Len(Text), Format, .RC, DTTO
                                        Else
                                            DrawThemeText CoolBarTheme, .hDC, RP_BAND, 0, StrPtr(Text), Len(Text), Format, 0, .RC
                                        End If
                                        SetBkMode .hDC, OldBkMode
                                        If hFontOld <> 0 Then SelectObject .hDC, hFontOld
                                    End If
                                End If
                                If (RBBI.fStyle And RBBS_USECHEVRON) = RBBS_USECHEVRON Then
                                    If ComCtlsSupportLevel() >= 2 Then
                                        Dim RBBI_V61 As REBARBANDINFO_V61
                                        RBBI_V61.RBBI.cbSize = LenB(RBBI_V61)
                                        RBBI_V61.RBBI.fMask = RBBIM_CHEVRONLOCATION Or RBBIM_CHEVRONSTATE
                                        SendMessage CoolBarHandle, RB_GETBANDINFO, Index, ByVal VarPtr(RBBI_V61)
                                        If (RBBI_V61.RCChevronLocation.Right - RBBI_V61.RCChevronLocation.Left) > 0 And (RBBI_V61.RCChevronLocation.Bottom - RBBI_V61.RCChevronLocation.Top) > 0 Then
                                            Const STATE_SYSTEM_PRESSED As Long = &H8, STATE_SYSTEM_HOTTRACKED As Long = &H80
                                            Dim ChevronPart As Long, ChevronState As Long
                                            If (dwStyle And CCS_VERT) = CCS_VERT Then
                                                ChevronPart = RP_CHEVRONVERT
                                            Else
                                                ChevronPart = RP_CHEVRON
                                            End If
                                            If (RBBI_V61.uChevronState And STATE_SYSTEM_PRESSED) = STATE_SYSTEM_PRESSED Then
                                                ChevronState = CHEVS_PRESSED
                                            ElseIf (RBBI_V61.uChevronState And STATE_SYSTEM_HOTTRACKED) = STATE_SYSTEM_HOTTRACKED Then
                                                ChevronState = CHEVS_HOT
                                            Else
                                                ChevronState = CHEVS_NORMAL
                                            End If
                                            If IsThemeBackgroundPartiallyTransparent(CoolBarTheme, ChevronPart, ChevronState) <> 0 Then DrawThemeParentBackground CoolBarHandle, .hDC, RBBI_V61.RCChevronLocation
                                            DrawThemeBackground CoolBarTheme, .hDC, ChevronPart, ChevronState, RBBI_V61.RCChevronLocation, RBBI_V61.RCChevronLocation
                                        End If
                                    End If
                                End If
                                End With
                                WindowProcUserControl = CDRF_SKIPDEFAULT
                            Else
                                WindowProcUserControl = CDRF_DODEFAULT
                            End If
                            Exit Function
                        
                        #End If
                        
                    End Select
                Case RBN_HEIGHTCHANGE
                    Call UserControl_Resize
                    RaiseEvent HeightChanged(Extender.Height)
                Case RBN_LAYOUTCHANGED
                    RaiseEvent LayoutChanged
                Case RBN_CHILDSIZE
                    Dim NMRBCS As NMREBARCHILDSIZE
                    CopyMemory NMRBCS, ByVal lParam, LenB(NMRBCS)
                    If NMRBCS.uBand > -1 Then
                        RBBI.cbSize = LenB(RBBI)
                        RBBI.fMask = RBBIM_LPARAM Or RBBIM_CHILDSIZE
                        SendMessage CoolBarHandle, RB_GETBANDINFO, NMRBCS.uBand, ByVal VarPtr(RBBI)
                        If RBBI.lParam <> 0 Then
                            Set Band = PtrToObj(RBBI.lParam)
                            If Not Band.Child Is Nothing Then
                                With Band.Child
                                .Move UserControl.ScaleX(NMRBCS.RCChild.Left, vbPixels, vbTwips), UserControl.ScaleY(NMRBCS.RCChild.Top, vbPixels, vbTwips), UserControl.ScaleX((NMRBCS.RCChild.Right - NMRBCS.RCChild.Left), vbPixels, vbTwips), UserControl.ScaleY((NMRBCS.RCChild.Bottom - NMRBCS.RCChild.Top), vbPixels, vbTwips)
                                Dim CY As Long
                                If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
                                    CY = UserControl.ScaleY(.Height, vbTwips, vbPixels)
                                Else
                                    CY = UserControl.ScaleX(.Width, vbTwips, vbPixels)
                                End If
                                If RBBI.CYMinChild < CY Then
                                    RBBI.fMask = RBBIM_CHILDSIZE
                                    RBBI.CYMinChild = CY
                                    SendMessage CoolBarHandle, RB_SETBANDINFO, NMRBCS.uBand, ByVal VarPtr(RBBI)
                                End If
                                End With
                            End If
                        End If
                    End If
                Case RBN_DELETINGBAND
                    CopyMemory NMRB, ByVal lParam, LenB(NMRB)
                    If NMRB.uBand > -1 Then
                        With RBBI
                        .cbSize = LenB(RBBI)
                        .fMask = RBBIM_CHILD
                        SendMessage CoolBarHandle, RB_GETBANDINFO, NMRB.uBand, ByVal VarPtr(RBBI)
                        If .hWndChild <> 0 And Not .hWndChild = -1 Then SetParent .hWndChild, UserControl.hWnd
                        End With
                    End If
                Case RBN_DELETEDBAND
                    CopyMemory NMRB, ByVal lParam, LenB(NMRB)
                    If NMRB.uBand = 0 Then
                        If SendMessage(CoolBarHandle, RB_GETBANDCOUNT, 0, ByVal 0&) > 0 Then
                            With RBBI
                            .cbSize = LenB(RBBI)
                            .fMask = RBBIM_STYLE
                            SendMessage CoolBarHandle, RB_GETBANDINFO, 0, ByVal VarPtr(RBBI)
                            If (.fStyle And RBBS_BREAK) = RBBS_BREAK Then
                                .fStyle = .fStyle And Not RBBS_BREAK
                                SendMessage CoolBarHandle, RB_SETBANDINFO, 0, ByVal VarPtr(RBBI)
                            End If
                            End With
                        End If
                    End If
                Case RBN_MINMAX
                    RaiseEvent MinMax(Cancel)
                    If Cancel = True Then
                        WindowProcUserControl = 1
                    Else
                        WindowProcUserControl = 0
                    End If
                    Exit Function
                Case RBN_BEGINDRAG
                    CopyMemory NMRB, ByVal lParam, LenB(NMRB)
                    If NMRB.uBand > -1 Then
                        If (NMRB.dwMask And RBNM_LPARAM) = RBNM_LPARAM Then
                            If NMRB.lParam <> 0 Then Set Band = PtrToObj(NMRB.lParam)
                        Else
                            RBBI.cbSize = LenB(RBBI)
                            RBBI.fMask = RBBIM_LPARAM
                            SendMessage CoolBarHandle, RB_GETBANDINFO, NMRB.uBand, ByVal VarPtr(RBBI)
                            If RBBI.lParam <> 0 Then Set Band = PtrToObj(RBBI.lParam)
                        End If
                        RaiseEvent BandBeforeDrag(Band, Cancel)
                        If Cancel = True Then
                            WindowProcUserControl = 1
                            Exit Function
                        End If
                    End If
                Case RBN_ENDDRAG
                    CopyMemory NMRB, ByVal lParam, LenB(NMRB)
                    If NMRB.uBand > -1 Then
                        If (NMRB.dwMask And RBNM_LPARAM) = RBNM_LPARAM Then
                            If NMRB.lParam <> 0 Then Set Band = PtrToObj(NMRB.lParam)
                        Else
                            RBBI.cbSize = LenB(RBBI)
                            RBBI.fMask = RBBIM_LPARAM
                            SendMessage CoolBarHandle, RB_GETBANDINFO, NMRB.uBand, ByVal VarPtr(RBBI)
                            If RBBI.lParam <> 0 Then Set Band = PtrToObj(RBBI.lParam)
                        End If
                        RaiseEvent BandAfterDrag(Band, NMRB.uBand + 1)
                    End If
                Case RBN_CHEVRONPUSHED
                    Dim NMRBCHEVR As NMREBARCHEVRON
                    CopyMemory NMRBCHEVR, ByVal lParam, LenB(NMRBCHEVR)
                    With NMRBCHEVR
                    If .uBand > -1 And .lParam <> 0 Then
                        Set Band = PtrToObj(.lParam)
                        With .RCChevron
                        RaiseEvent BandChevronPushed(Band, UserControl.ScaleX(.Left, vbPixels, vbContainerPosition), UserControl.ScaleY(.Top, vbPixels, vbContainerPosition), UserControl.ScaleX((.Right - .Left), vbPixels, vbContainerSize), UserControl.ScaleY((.Bottom - .Top), vbPixels, vbContainerSize))
                        End With
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
End Function

Private Function WindowProcUserControlDesignMode(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR, RBBI As REBARBANDINFO
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = CoolBarHandle Then
            Select Case NM.Code
                Case NM_CUSTOMDRAW
                    WindowProcUserControlDesignMode = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
                    Exit Function
                Case RBN_HEIGHTCHANGE
                    Call UserControl_Resize
                Case RBN_CHILDSIZE
                    Dim NMRBCS As NMREBARCHILDSIZE
                    CopyMemory NMRBCS, ByVal lParam, LenB(NMRBCS)
                    If NMRBCS.uBand > -1 Then
                        Dim Band As CbrBand
                        RBBI.cbSize = LenB(RBBI)
                        RBBI.fMask = RBBIM_LPARAM Or RBBIM_CHILDSIZE
                        SendMessage CoolBarHandle, RB_GETBANDINFO, NMRBCS.uBand, ByVal VarPtr(RBBI)
                        If RBBI.lParam <> 0 Then
                            Set Band = PtrToObj(RBBI.lParam)
                            If Not Band.Child Is Nothing Then
                                Dim WndRect As RECT, ClientRect As RECT
                                GetWindowRect CoolBarHandle, WndRect
                                GetClientRect CoolBarHandle, ClientRect
                                With Band.Child
                                If PropRightToLeft = True And PropRightToLeftLayout = True Then
                                    MapWindowPoints CoolBarHandle, UserControl.hWnd, NMRBCS.RCChild, 2
                                    With NMRBCS.RCChild
                                    .Left = .Left - (((WndRect.Right - WndRect.Left) - (ClientRect.Right - ClientRect.Left)) / 2)
                                    .Top = .Top - (((WndRect.Right - WndRect.Left) - (ClientRect.Right - ClientRect.Left)) / 2)
                                    .Right = .Right - (((WndRect.Right - WndRect.Left) - (ClientRect.Right - ClientRect.Left)) / 2)
                                    .Bottom = .Bottom - (((WndRect.Right - WndRect.Left) - (ClientRect.Right - ClientRect.Left)) / 2)
                                    End With
                                End If
                                .Move UserControl.ScaleX(NMRBCS.RCChild.Left + (((WndRect.Right - WndRect.Left) - (ClientRect.Right - ClientRect.Left)) / 2), vbPixels, vbTwips), UserControl.ScaleY(NMRBCS.RCChild.Top + (((WndRect.Bottom - WndRect.Top) - (ClientRect.Bottom - ClientRect.Top)) / 2), vbPixels, vbTwips), UserControl.ScaleX((NMRBCS.RCChild.Right - NMRBCS.RCChild.Left), vbPixels, vbTwips), UserControl.ScaleY((NMRBCS.RCChild.Bottom - NMRBCS.RCChild.Top), vbPixels, vbTwips)
                                Dim CY As Long
                                If (GetWindowLong(CoolBarHandle, GWL_STYLE) And CCS_VERT) = 0 Then
                                    CY = UserControl.ScaleY(.Height, vbTwips, vbPixels)
                                Else
                                    CY = UserControl.ScaleX(.Width, vbTwips, vbPixels)
                                End If
                                If RBBI.CYMinChild < CY Then
                                    RBBI.fMask = RBBIM_CHILDSIZE
                                    RBBI.CYMinChild = CY
                                    SendMessage CoolBarHandle, RB_SETBANDINFO, NMRBCS.uBand, ByVal VarPtr(RBBI)
                                End If
                                End With
                            End If
                        End If
                    End If
                Case RBN_DELETEDBAND
                    Dim NMRB As NMREBAR
                    CopyMemory NMRB, ByVal lParam, LenB(NMRB)
                    If NMRB.uBand = 0 Then
                        If SendMessage(CoolBarHandle, RB_GETBANDCOUNT, 0, ByVal 0&) > 0 Then
                            With RBBI
                            .cbSize = LenB(RBBI)
                            .fMask = RBBIM_STYLE
                            SendMessage CoolBarHandle, RB_GETBANDINFO, 0, ByVal VarPtr(RBBI)
                            If (.fStyle And RBBS_BREAK) = RBBS_BREAK Then
                                .fStyle = .fStyle And Not RBBS_BREAK
                                SendMessage CoolBarHandle, RB_SETBANDINFO, 0, ByVal VarPtr(RBBI)
                            End If
                            End With
                        End If
                    End If
            End Select
        End If
End Select
WindowProcUserControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
