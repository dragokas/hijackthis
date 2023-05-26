VERSION 5.00
Begin VB.UserControl RichTextBox 
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H80000008&
   HasDC           =   0   'False
   PropertyPages   =   "RichTextBox.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "RichTextBox.ctx":004C
End
Attribute VB_Name = "RichTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const ImplementThemedBorder = True

#If False Then
Private RtfLoadSaveFormatRTF, RtfLoadSaveFormatText, RtfLoadSaveFormatUnicodeText
Private RtfFindOptionWholeWord, RtfFindOptionMatchCase, RtfFindOptionNoHighlight, RtfFindOptionReverse
Private RtfActionTypeUnknown, RtfActionTypeTyping, RtfActionTypeDelete, RtfActionTypeOLEDragDrop, RtfActionTypeCut, RtfActionTypePaste, RtfActionTypeAutoTable
Private RtfSelAlignmentLeft, RtfSelAlignmentRight, RtfSelAlignmentCenter, RtfSelAlignmentJustified
Private RtfSelTypeEmpty, RtfSelTypeText, RtfSelTypeObject, RtfSelTypeMultiChar, RtfSelTypeMultiObject
Private RtfTextModeRichText, RtfTextModePlainText
#End If
Public Enum RtfLoadSaveFormatConstants
RtfLoadSaveFormatRTF = 0
RtfLoadSaveFormatText = 1
RtfLoadSaveFormatUnicodeText = 2
End Enum
Private Const FR_WHOLEWORD As Long = &H2
Private Const FR_MATCHCASE As Long = &H4
Public Enum RtfFindOptionConstants
RtfFindOptionWholeWord = FR_WHOLEWORD
RtfFindOptionMatchCase = FR_MATCHCASE
RtfFindOptionNoHighlight = &H8
RtfFindOptionReverse = &H10
End Enum
Private Const UID_UNKNOWN As Long = 0
Private Const UID_TYPING As Long = 1
Private Const UID_DELETE As Long = 2
Private Const UID_DRAGDROP As Long = 3
Private Const UID_CUT As Long = 4
Private Const UID_PASTE As Long = 5
Private Const UID_AUTOTABLE As Long = 6
Public Enum RtfActionTypeConstants
RtfActionTypeUnknown = UID_UNKNOWN
RtfActionTypeTyping = UID_TYPING
RtfActionTypeDelete = UID_DELETE
RtfActionTypeOLEDragDrop = UID_DRAGDROP
RtfActionTypeCut = UID_CUT
RtfActionTypePaste = UID_PASTE
RtfActionTypeAutoTable = UID_AUTOTABLE
End Enum
Public Enum RtfSelAlignmentConstants
RtfSelAlignmentLeft = 0
RtfSelAlignmentRight = 1
RtfSelAlignmentCenter = 2
RtfSelAlignmentJustified = 3
End Enum
Private Const SEL_EMPTY As Long = 0
Private Const SEL_TEXT As Long = 1
Private Const SEL_OBJECT As Long = 2
Private Const SEL_MULTICHAR As Long = 4
Private Const SEL_MULTIOBJECT As Long = 8
Public Enum RtfSelTypeConstants
RtfSelTypeEmpty = SEL_EMPTY
RtfSelTypeText = SEL_TEXT
RtfSelTypeObject = SEL_OBJECT
RtfSelTypeMultiChar = SEL_MULTICHAR
RtfSelTypeMultiObject = SEL_MULTIOBJECT
End Enum
Public Enum RtfTextModeConstants
RtfTextModeRichText = 0
RtfTextModePlainText = 1
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
Private Type DOCINFO
cbSize As Long
lpszDocName As Long
lpszOutput As Long
lpszDatatype As Long
fwType As Long
End Type
Private Const LF_FACESIZE As Long = 32
Private Type RECHARFORMAT2
cbSize As Long
dwMask As Long
dwEffects As Long
YHeight As Long
YOffset As Long
TextColor As Long
Charset As Byte
PitchAndFamily As Byte
FaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
Weight As Integer
Spacing As Integer
BackColor As Long
LCID As Long
dwReserved As Long
Style As Integer
Kerning As Integer
UnderlineType As Byte
Animation As Byte
RevAuthor As Byte
UnderlineColor As Byte
End Type
Private Const MAX_TAB_STOPS As Long = 32
Private Type REPARAFORMAT2
cbSize As Long
dwMask As Long
Numbering As Integer
Effects As Integer
DXStartIndent As Long
DXRightIndent As Long
DXOffset As Long
Alignment As Integer
TabCount As Integer
Tabs(0 To (MAX_TAB_STOPS - 1)) As Long
DYSpaceBefore As Long
DYSpaceAfter As Long
DYLineSpacing As Long
Style As Integer
LineSpacingRule As Byte
OutlineLevel As Byte
ShadingWeight As Integer
ShadingStyle As Integer
NumberingStart As Integer
NumberingStyle As Integer
NumberingTab As Integer
BorderSpace As Integer
BorderWidth As Integer
Borders As Integer
End Type
Private Type REGETTEXTLENGTHEX
Flags As Long
CodePage As Long
End Type
Private Type REGETTEXTEX
cbSize As Long
Flags As Long
CodePage As Long
lpDefaultChar As Long
lpUsedDefChar As Long
End Type
Private Type RESETTEXTEX
Flags As Long
CodePage As Long
End Type
Private Type RECHARRANGE
Min As Long
Max As Long
End Type
Private Type RETEXTRANGE
CharRange As RECHARRANGE
lpstrText As Long
End Type
Private Type REFINDTEXTEX
CharRange As RECHARRANGE
lpstrText As Long
CharRangeText As RECHARRANGE
End Type
Private Type REEDITSTREAM
dwCookie As Long
dwError As Long
lpfnCallback As Long
End Type
Private Type REOBJECT
cbStruct As Long
CharPos As Long
riid As OLEGuids.OLECLSID
pOleObject As OLEGuids.IOleObject
pStorage As OLEGuids.IStorage
pOleSite As OLEGuids.IOleClientSite
Size As SIZEAPI
dvAspect As Long
dwFlags As Long
dwUser As Long
End Type
Private Type REFORMATRANGE
hDC As Long
hDCTarget As Long
RC As RECT
RCPage As RECT
CharRange As RECHARRANGE
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMENSELCHANGE
hdr As NMHDR
CharRange As RECHARRANGE
SelType As Integer
End Type
Private Type NMENLINK
hdr As NMHDR
wMsg As Long
wParam As Long
lParam As Long
CharRange As RECHARRANGE
End Type
Private Type NMENPROTECTED
hdr As NMHDR
wMsg As Long
wParam As Long
lParam As Long
CharRange As RECHARRANGE
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
Public Event SelChange(ByVal SelType As Integer, ByVal SelStart As Long, ByVal SelEnd As Long)
Attribute SelChange.VB_Description = "Occurs when the current selection of text in a control has changed or the insertion point has moved."
Public Event LinkEvent(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal LinkStart As Long, ByVal LinkEnd As Long)
Attribute LinkEvent.VB_Description = "Occurs on various reasons, for example, when the user clicks the mouse or when the mouse pointer is over text that has a link format."
Public Event ModifyProtected(ByRef Allow As Boolean, ByVal SelStart As Long, ByVal SelEnd As Long)
Attribute ModifyProtected.VB_Description = "Occurs when the user attempts to edit protected text."
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
Public Event OLECompleteDrag()
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a drag/drop has been completed or canceled."
Public Event OLEGetDropEffect(ByRef Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute OLEGetDropEffect.VB_Description = "Occurs during an OLE drag/drop operation to specify the effect of which indicates what the result of the drop operation would be."
Public Event OLEStartDrag(ByRef AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated."
Public Event OLEGetContextMenu(ByVal SelType As Integer, ByVal LpOleObject As Long, ByVal SelStart As Long, ByVal SelEnd As Long, ByRef hMenu As Long)
Attribute OLEGetContextMenu.VB_Description = "This is a request to provide a popup menu to use on a right-click. The rich text box control destroys the popup menu when it is finished."
Public Event OLEContextMenuClick(ByVal ID As Long)
Attribute OLEContextMenuClick.VB_Description = "Occurs when the user selects an item from a popup menu that was provided in the OLEGetContextMenu event."
Public Event OLEDeleteObject(ByVal LpOleObject As Long)
Attribute OLEDeleteObject.VB_Description = "Occurs when an OLE object is about to be deleted. The OLE object is not necessarily being released."
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
Private Declare Function StgCreateDocFile Lib "ole32" Alias "StgCreateDocfile" (ByVal pwcsName As Long, ByVal grfMode As Long, ByVal Reserved As Long, ByRef ppStgOpen As OLEGuids.IStorage) As Long
Private Declare Function OleSetContainedObject Lib "ole32" (ByVal pUnknown As IUnknown, ByVal fContained As Long) As Long
Private Declare Function OleCreateFromFile Lib "ole32" (ByRef pCLSID As Any, ByVal lpszFileName As Long, ByRef riid As Any, ByVal RenderOpt As Long, ByVal lpFormatEtc As Long, ByVal pClientSite As OLEGuids.IOleClientSite, ByVal pStg As OLEGuids.IStorage, ByRef ppvObj As OLEGuids.IOleObject) As Long
Private Declare Function OleCreateLinkToFile Lib "ole32" (ByVal lpszFileName As Long, ByRef riid As Any, ByVal RenderOpt As Long, ByVal lpFormatEtc As Long, ByVal pClientSite As OLEGuids.IOleClientSite, ByVal pStg As OLEGuids.IStorage, ByRef ppvObj As OLEGuids.IOleObject) As Long
Private Declare Function OleQueryCreateFromData Lib "ole32" (ByVal pSrcDataObject As OLEGuids.IDataObject) As Long
Private Declare Function OleCreateStaticFromData Lib "ole32" (ByVal pSrcDataObject As OLEGuids.IDataObject, ByRef riid As Any, ByVal RenderOpt As Long, ByVal lpFormatEtc As Long, ByVal pClientSite As OLEGuids.IOleClientSite, ByVal pStg As OLEGuids.IStorage, ByRef ppvObj As OLEGuids.IOleObject) As Long
Private Declare Function OleCreateFromData Lib "ole32" (ByVal pSrcDataObject As OLEGuids.IDataObject, ByRef riid As Any, ByVal RenderOpt As Long, ByVal lpFormatEtc As Long, ByVal pClientSite As OLEGuids.IOleClientSite, ByVal pStg As OLEGuids.IStorage, ByRef ppvObj As OLEGuids.IOleObject) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToWrite As Long, ByRef NumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToRead As Long, ByRef NumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function StartDoc Lib "gdi32" Alias "StartDocW" (ByVal hDC As Long, ByRef lpDI As DOCINFO) As Long
Private Declare Function EndDoc Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function AbortDoc Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatW" (ByVal lpString As Long) As Long

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
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const CF_UNICODETEXT As Long = 13
Private Const CP_UNICODE As Long = 1200
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const SB_THUMBTRACK As Long = 5
Private Const SW_HIDE As Long = &H0
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_ENABLE As Long = &HA
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_STYLECHANGED As Long = &H7D
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
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_SETTEXT As Long = &HC
Private Const WM_COPY As Long = &H301
Private Const WM_CUT As Long = &H300
Private Const WM_PASTE As Long = &H302
Private Const WM_CLEAR As Long = &H303
Private Const WM_NCPAINT As Long = &H85
Private Const WM_PRINT As Long = &H317, PRF_NONCLIENT As Long = &H2
Private Const EM_SETREADONLY As Long = &HCF, ES_READONLY As Long = &H800
Private Const EM_SCROLL As Long = &HB5
Private Const EM_LINESCROLL As Long = &HB6
Private Const EM_SCROLLCARET As Long = &HB7
Private Const EM_REPLACESEL As Long = &HC2
Private Const EM_GETPASSWORDCHAR As Long = &HD2
Private Const EM_SETPASSWORDCHAR As Long = &HCC
Private Const EM_GETMODIFY As Long = &HB8
Private Const EM_SETMODIFY As Long = &HB9
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_GETTHUMB As Long = &HBE
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_GETLINE As Long = &HC4
Private Const EM_UNDO As Long = &HC7
Private Const EM_CANUNDO As Long = &HC6
Private Const EM_EMPTYUNDOBUFFER As Long = &HCD
Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_GETMARGINS As Long = &HD4
Private Const EM_SETMARGINS As Long = &HD3
Private Const EM_POSFROMCHAR As Long = &HD6
Private Const EM_CHARFROMPOS As Long = &HD7
Private Const WM_USER As Long = &H400
Private Const EM_CANPASTE As Long = (WM_USER + 50)
Private Const EM_DISPLAYBAND As Long = (WM_USER + 51)
Private Const EM_EXGETSEL As Long = (WM_USER + 52)
Private Const EM_EXLIMITTEXT As Long = (WM_USER + 53)
Private Const EM_EXLINEFROMCHAR As Long = (WM_USER + 54)
Private Const EM_EXSETSEL As Long = (WM_USER + 55)
Private Const EM_FINDTEXTA As Long = (WM_USER + 56)
Private Const EM_FINDTEXTW As Long = (WM_USER + 123)
Private Const EM_FINDTEXT As Long = EM_FINDTEXTW
Private Const EM_FINDTEXTEXA As Long = (WM_USER + 79)
Private Const EM_FINDTEXTEXW As Long = (WM_USER + 124)
Private Const EM_FINDTEXTEX As Long = EM_FINDTEXTEXW
Private Const EM_FORMATRANGE As Long = (WM_USER + 57)
Private Const EM_GETCHARFORMAT As Long = (WM_USER + 58)
Private Const EM_GETEVENTMASK As Long = (WM_USER + 59)
Private Const EM_GETOLEINTERFACE As Long = (WM_USER + 60)
Private Const EM_GETPARAFORMAT As Long = (WM_USER + 61)
Private Const EM_GETSELTEXT As Long = (WM_USER + 62)
Private Const EM_HIDESELECTION As Long = (WM_USER + 63)
Private Const EM_PASTESPECIAL As Long = (WM_USER + 64)
Private Const EM_SELECTIONTYPE As Long = (WM_USER + 66)
Private Const EM_SETBKGNDCOLOR As Long = (WM_USER + 67)
Private Const EM_SETCHARFORMAT As Long = (WM_USER + 68)
Private Const EM_SETEVENTMASK As Long = (WM_USER + 69)
Private Const EM_SETOLECALLBACK As Long = (WM_USER + 70)
Private Const EM_SETPARAFORMAT As Long = (WM_USER + 71)
Private Const EM_SETTARGETDEVICE As Long = (WM_USER + 72)
Private Const EM_STREAMIN As Long = (WM_USER + 73)
Private Const EM_STREAMOUT As Long = (WM_USER + 74)
Private Const EM_GETTEXTRANGE As Long = (WM_USER + 75)
Private Const EM_SETOPTIONS As Long = (WM_USER + 77)
Private Const EM_GETOPTIONS As Long = (WM_USER + 78)
Private Const EM_SETUNDOLIMIT As Long = (WM_USER + 82)
Private Const EM_REDO As Long = (WM_USER + 84)
Private Const EM_CANREDO As Long = (WM_USER + 85)
Private Const EM_GETUNDONAME As Long = (WM_USER + 86)
Private Const EM_GETREDONAME As Long = (WM_USER + 87)
Private Const EM_STOPGROUPTYPING As Long = (WM_USER + 88)
Private Const EM_SETTEXTMODE As Long = (WM_USER + 89)
Private Const EM_GETTEXTMODE As Long = (WM_USER + 90)
Private Const EM_AUTOURLDETECT As Long = (WM_USER + 91)
Private Const EM_GETAUTOURLDETECT As Long = (WM_USER + 92)
Private Const EM_GETTEXTEX As Long = (WM_USER + 94)
Private Const EM_GETTEXTLENGTHEX As Long = (WM_USER + 95)
Private Const EM_SETTEXTEX As Long = (WM_USER + 97)
Private Const EM_SETLANGOPTIONS As Long = (WM_USER + 120)
Private Const EM_GETLANGOPTIONS As Long = (WM_USER + 121)
Private Const EM_SETTYPOGRAPHYOPTIONS As Long = (WM_USER + 202)
Private Const EM_SETEDITSTYLE As Long = (WM_USER + 204)
Private Const EM_GETEDITSTYLE As Long = (WM_USER + 205)
Private Const EM_GETSCROLLPOS As Long = (WM_USER + 221)
Private Const EM_SETSCROLLPOS As Long = (WM_USER + 222)
Private Const EM_GETZOOM As Long = (WM_USER + 224)
Private Const EM_SETZOOM As Long = (WM_USER + 225)
Private Const EM_SETEDITSTYLEEX As Long = (WM_USER + 275)
Private Const EM_GETEDITSTYLEEX As Long = (WM_USER + 276)
Private Const ENM_NONE As Long = &H0
Private Const ENM_CHANGE As Long = &H1
Private Const ENM_SCROLL As Long = &H4
Private Const ENM_KEYEVENTS As Long = &H10000
Private Const ENM_MOUSEEVENTS As Long = &H20000
Private Const ENM_SELCHANGE As Long = &H80000
Private Const ENM_DROPFILES As Long = &H100000 ' Not applicable if an IRichEditOleCallback is set.
Private Const ENM_PROTECTED As Long = &H200000
Private Const ENM_CORRECTTEXT As Long = &H400000
Private Const ENM_SCROLLEVENTS As Long = &H8
Private Const ENM_DRAGDROPDONE As Long = &H10
Private Const ENM_IMECHANGE As Long = &H800000
Private Const ENM_LANGCHANGE As Long = &H1000000
Private Const ENM_LINK As Long = &H4000000
Private Const EN_CHANGE As Long = &H300
Private Const EN_MAXTEXT As Long = &H501
Private Const EN_HSCROLL As Long = &H601
Private Const EN_VSCROLL As Long = &H602
Private Const EN_SELCHANGE As Long = &H702
Private Const EN_DROPFILES As Long = &H703 ' Not applicable if an IRichEditOleCallback is set.
Private Const EN_PROTECTED As Long = &H704
Private Const EN_SAVECLIPBOARD As Long = &H708
Private Const EN_LINK As Long = &H70B
Private Const EN_DRAGDROPDONE As Long = &H70C
Private Const ES_AUTOHSCROLL As Long = &H80
Private Const ES_AUTOVSCROLL As Long = &H40
Private Const ES_NOHIDESEL As Long = &H100
Private Const ES_MULTILINE As Long = &H4
Private Const ES_NOOLEDRAGDROP As Long = &H8
Private Const ES_PASSWORD As Long = &H20
Private Const ES_WANTRETURN As Long = &H1000
Private Const ES_DISABLENOSCROLL As Long = &H2000
Private Const ES_SUNKEN As Long = &H4000
Private Const ES_SAVESEL As Long = &H8000& ' Malfunction
Private Const ES_SELECTIONBAR As Long = &H1000000
Private Const ES_VERTICAL As Long = &H400000
Private Const EC_LEFTMARGIN As Long = &H1
Private Const EC_RIGHTMARGIN As Long = &H2
Private Const EC_USEFONTINFO As Long = &HFFFF&
Private Const SES_EMULATESYSEDIT As Long = 1
Private Const SES_BEEPONMAXTEXT As Long = 2
Private Const SES_EXTENDBACKCOLOR As Long = 4
Private Const SES_MAPCPS As Long = 8 ' Obsolete
Private Const SES_EMULATE10 As Long = 16
Private Const SES_USECRLF As Long = 32 ' Obsolete
Private Const SES_USEAIMM As Long = 64
Private Const SES_NOIME As Long = 128
Private Const SES_ALLOWBEEPS As Long = 256
Private Const SES_UPPERCASE As Long = 512
Private Const SES_LOWERCASE As Long = 1024
Private Const SES_NOINPUTSEQUENCECHK As Long = 2048
Private Const SES_BIDI As Long = 4096
Private Const SES_SCROLLONKILLFOCUS As Long = 8192
Private Const SES_XLTCRCRLFTOCR As Long = 16384
Private Const GTL_DEFAULT As Long = 0
Private Const GTL_USECRLF As Long = 1
Private Const GTL_PRECISE As Long = 2
Private Const GTL_CLOSE As Long = 4
Private Const GTL_NUMCHARS As Long = 8
Private Const GTL_NUMBYTES As Long = 16
Private Const GT_DEFAULT As Long = 0
Private Const GT_USECRLF As Long = 1
Private Const GT_SELECTION As Long = 2
Private Const GT_RAWTEXT As Long = 4
Private Const GT_NOHIDDENTEXT As Long = 8
Private Const ST_DEFAULT As Long = 0
Private Const ST_KEEPUNDO As Long = 1
Private Const ST_SELECTION As Long = 2
Private Const ST_NEWCHARS As Long = 4
Private Const ST_UNICODE As Long = 8
Private Const SF_TEXT As Long = &H1
Private Const SF_RTF As Long = &H2
Private Const SF_RTFNOOBJS As Long = &H3
Private Const SF_TEXTIZED As Long = &H4
Private Const SF_UNICODE As Long = &H10
Private Const SF_USECODEPAGE As Long = &H20
Private Const SFF_SELECTION As Long = &H8000&
Private Const SFF_PLAINRTF As Long = &H4000
Private Const SCF_DEFAULT As Long = &H0
Private Const SCF_SELECTION As Long = &H1
Private Const CFM_BOLD As Long = &H1
Private Const CFM_ITALIC As Long = &H2
Private Const CFM_UNDERLINE As Long = &H4
Private Const CFM_STRIKEOUT As Long = &H8
Private Const CFM_PROTECTED As Long = &H10
Private Const CFM_HIDDEN As Long = &H100
Private Const CFM_LINK As Long = &H20
Private Const CFM_SIZE As Long = &H80000000
Private Const CFM_COLOR As Long = &H40000000
Private Const CFM_FACE As Long = &H20000000
Private Const CFM_OFFSET As Long = &H10000000
Private Const CFM_BACKCOLOR As Long = &H4000000
Private Const CFM_CHARSET As Long = &H8000000
Private Const CFE_BOLD As Long = &H1
Private Const CFE_ITALIC As Long = &H2
Private Const CFE_UNDERLINE As Long = &H4
Private Const CFE_STRIKEOUT As Long = &H8
Private Const CFE_PROTECTED As Long = &H10
Private Const CFE_HIDDEN As Long = &H100
Private Const CFE_LINK As Long = &H20
Private Const CFE_SUBSCRIPT As Long = &H10000
Private Const CFE_SUPERSCRIPT As Long = &H20000
Private Const CFE_AUTOCOLOR As Long = &H40000000
Private Const PFM_NUMBERING As Long = &H20
Private Const PFM_ALIGNMENT As Long = &H8
Private Const PFM_SPACEBEFORE As Long = &H40
Private Const PFM_NUMBERINGSTYLE As Long = &H2000
Private Const PFM_NUMBERINGSTART As Long = &H8000&
Private Const PFM_BORDER As Long = &H800
Private Const PFM_RIGHTINDENT As Long = &H2
Private Const PFM_STARTINDENT As Long = &H1
Private Const PFM_OFFSET As Long = &H4
Private Const PFM_OFFSETINDENT As Long = &H80000000
Private Const PFM_LINESPACING As Long = &H100
Private Const PFM_SPACEAFTER As Long = &H80
Private Const PFM_NUMBERINGTAB As Long = &H4000
Private Const PFM_TABLE As Long = &H40000000
Private Const PFM_TABSTOPS As Long = &H10
Private Const PFA_LEFT As Long = 1
Private Const PFA_RIGHT As Long = 2
Private Const PFA_CENTER As Long = 3
Private Const PFA_JUSTIFY As Long = 4
Private Const PFN_BULLET As Long = 1
Private Const TO_ADVANCEDTYPOGRAPHY As Long = 1
Private Const TM_PLAINTEXT As Long = 1
Private Const TM_RICHTEXT As Long = 2
Private Const TM_SINGLELEVELUNDO As Long = 4
Private Const TM_MULTILEVELUNDO As Long = 8
Private Const ECO_AUTOWORDSELECTION As Long = 1
Private Const ECO_AUTOVSCROLL As Long = ES_AUTOVSCROLL
Private Const ECO_AUTOHSCROLL As Long = ES_AUTOHSCROLL
Private Const ECO_NOHIDESEL As Long = ES_NOHIDESEL
Private Const ECO_READONLY As Long = ES_READONLY
Private Const ECO_WANTRETURN As Long = ES_WANTRETURN
Private Const ECO_SAVESEL As Long = ES_SAVESEL
Private Const ECO_SELECTIONBAR As Long = ES_SELECTIONBAR
Private Const ECO_VERTICAL As Long = ES_VERTICAL
Private Const ECOOP_SET As Long = 1
Private Const ECOOP_OR As Long = 2
Private Const ECOOP_AND As Long = 3
Private Const ECOOP_XOR As Long = 4
Private Const STGM_CREATE As Long = &H0
Private Const STGM_READWRITE As Long = &H2
Private Const STGM_SHARE_EXCLUSIVE As Long = &H10
Private Const STGM_DELETEONRELEASE As Long = &H4000000
Private Const REO_GETOBJ_NO_INTERFACES As Long = 0
Private Const REO_GETOBJ_POLEOBJ As Long = 1
Private Const REO_GETOBJ_PSTG As Long = 2
Private Const REO_GETOBJ_POLESITE As Long = 4
Private Const REO_GETOBJ_ALL_INTERFACES As Long = 7
Private Const REO_IOB_SELECTION As Long = &HFFFFFFFF
Private Const REO_CP_SELECTION As Long = &HFFFFFFFF
Private Const REO_IOB_USE_CP As Long = &HFFFFFFFE
Private Const REO_NULL As Long = 0
Private Const REO_RESIZABLE As Long = 1
Private Const REO_BELOWBASELINE As Long = 2
Private Const REO_INVERTEDSELECT As Long = 4
Private Const REO_DYNAMICSIZE As Long = 8
Private Const REO_BLANK As Long = 16
Private Const REO_DONTNEEDPALETTE As Long = 32
Private Const REO_READWRITEMASK As Long = 63
Private Const REO_GETMETAFILE As Long = &H400000
Private Const REO_LINKAVAILABLE As Long = &H800000
Private Const REO_HILITED As Long = &H1000000
Private Const REO_INPLACEACTIVE As Long = &H2000000
Private Const REO_OPEN As Long = &H4000000
Private Const REO_SELECTED As Long = &H8000000
Private Const REO_STATIC As Long = &H40000000
Private Const REO_LINK As Long = &H80000000
Private Const OLE_S_STATIC As Long = &H40001
Private Const S_OK As Long = &H0
Private Const OLERENDER_DRAW As Long = 1
Private Const DVASPECT_CONTENT As Long = 1
Private Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Private Const INVALID_HANDLE_VALUE As Long = (-1)
Private Const CREATE_ALWAYS As Long = 2
Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_READ As Long = &H80000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_BEGIN As Long = 0
Private Const PHYSICALWIDTH As Long = 110
Private Const PHYSICALHEIGHT As Long = 111
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Const HORZRES As Long = 8
Private Const VERTRES As Long = 10
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IOleControlVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private RichTextBoxHandle As Long
Private RichTextBoxFontHandle As Long
Private RichTextBoxIMCHandle As Long
Private RichTextBoxCharCodeCache As Long
Private RichTextBoxIsClick As Boolean
Private RichTextBoxMouseOver(0 To 1) As Boolean
Private RichTextBoxDesignMode As Boolean
Private RichTextBoxFocused As Boolean
Private RichTextBoxDataObjectValue As Variant
Private RichTextBoxDataObjectFormat As Variant
Private RichTextBoxIsOleCallback As Boolean
Private RichTextBoxEnabledVisualStyles As Boolean
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private DispIDBorderStyle As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropOLEDragDrop As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBorderStyle As Integer
Private PropBackColor As OLE_COLOR
Private PropLocked As Boolean
Private PropHideSelection As Boolean
Private PropPasswordChar As Integer
Private PropUseSystemPasswordChar As Boolean
Private PropMultiLine As Boolean
Private PropMaxLength As Long
Private PropScrollBars As VBRUN.ScrollBarConstants
Private PropWantReturn As Boolean
Private PropDisableNoScroll As Boolean
Private PropAutoURLDetect As Boolean
Private PropBulletIndent As Long
Private PropSelectionBar As Boolean
Private PropFileName As String
Private PropTextMode As RtfTextModeConstants
Private PropUndoLimit As Long
Private PropIMEMode As CCIMEModeConstants
Private PropAllowOverType As Boolean
Private PropOverTypeMode As Boolean
Private PropUseCrLf As Boolean

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
ElseIf DispID = DispIDBorderStyle Then
    Select Case PropBorderStyle
        Case vbBSNone: DisplayName = vbBSNone & " - None"
        Case vbFixedSingle: DisplayName = vbFixedSingle & " - Fixed Single"
    End Select
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
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Then
    Value = Cookie
    Handled = True
ElseIf DispID = DispIDBorderStyle Then
    Value = Cookie
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call RtfLoadRichedMod
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfaceControl)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDBorderStyle = 0 Then DispIDBorderStyle = GetDispID(Me, "BorderStyle")
On Error Resume Next
RichTextBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropOLEDragDrop = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = vbFixedSingle
PropBackColor = vbWindowBackground
PropLocked = False
PropHideSelection = True
PropPasswordChar = 0
PropUseSystemPasswordChar = False
PropMultiLine = True
PropMaxLength = 0
PropScrollBars = vbSBNone
PropWantReturn = True
PropDisableNoScroll = False
PropAutoURLDetect = True
PropBulletIndent = 0
PropSelectionBar = False
PropFileName = vbNullString
PropTextMode = RtfTextModeRichText
PropUndoLimit = 100
PropIMEMode = CCIMEModeNoControl
PropAllowOverType = True
PropOverTypeMode = False
Call CreateRichTextBox
Me.Text = Ambient.DisplayName
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDBorderStyle = 0 Then DispIDBorderStyle = GetDispID(Me, "BorderStyle")
On Error Resume Next
RichTextBoxDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
PropOLEDragDrop = .ReadProperty("OLEDragDrop", True)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropBorderStyle = .ReadProperty("BorderStyle", vbFixedSingle)
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
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
PropWantReturn = .ReadProperty("WantReturn", True)
PropDisableNoScroll = .ReadProperty("DisableNoScroll", False)
PropAutoURLDetect = .ReadProperty("AutoURLDetect", True)
PropBulletIndent = .ReadProperty("BulletIndent", 0)
PropSelectionBar = .ReadProperty("SelectionBar", False)
PropFileName = VarToStr(.ReadProperty("FileName", vbNullString))
PropTextMode = .ReadProperty("TextMode", RtfTextModeRichText)
PropUndoLimit = .ReadProperty("UndoLimit", 100)
PropIMEMode = .ReadProperty("IMEMode", CCIMEModeNoControl)
PropAllowOverType = .ReadProperty("AllowOverType", True)
PropOverTypeMode = .ReadProperty("OverTypeMode", False)
PropUseCrLf = .ReadProperty("UseCrLf", False)
End With
Call CreateRichTextBox
If PropTextMode = RtfTextModeRichText Then
    StreamStringIn VarToStr(PropBag.ReadProperty("TextRTF", vbNullString)), SF_RTF
Else
    StreamStringIn VarToStr(PropBag.ReadProperty("Text", vbNullString)), SF_TEXT Or SF_UNICODE
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragDrop", PropOLEDragDrop, True
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BorderStyle", PropBorderStyle, vbFixedSingle
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "Locked", PropLocked, False
.WriteProperty "HideSelection", PropHideSelection, True
.WriteProperty "PasswordChar", PropPasswordChar, 0
.WriteProperty "UseSystemPasswordChar", PropUseSystemPasswordChar, False
.WriteProperty "MultiLine", PropMultiLine, False
.WriteProperty "MaxLength", PropMaxLength, 0
.WriteProperty "ScrollBars", PropScrollBars, vbSBNone
.WriteProperty "WantReturn", PropWantReturn, True
.WriteProperty "DisableNoScroll", PropDisableNoScroll, False
.WriteProperty "AutoURLDetect", PropAutoURLDetect, True
.WriteProperty "BulletIndent", PropBulletIndent, 0
.WriteProperty "SelectionBar", PropSelectionBar, False
.WriteProperty "FileName", StrToVar(PropFileName), vbNullString
.WriteProperty "TextMode", PropTextMode, RtfTextModeRichText
.WriteProperty "UndoLimit", PropUndoLimit, 100
.WriteProperty "IMEMode", PropIMEMode, CCIMEModeNoControl
.WriteProperty "AllowOverType", PropAllowOverType, True
.WriteProperty "OverTypeMode", PropOverTypeMode, False
.WriteProperty "UseCrLf", PropUseCrLf, False
Dim Buffer As String
StreamStringOut Buffer, SF_TEXT Or SF_UNICODE
.WriteProperty "Text", StrToVar(Buffer), vbNullString
Buffer = vbNullString
StreamStringOut Buffer, SF_RTF
.WriteProperty "TextRTF", StrToVar(Buffer), vbNullString
End With
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
If Not IsEmpty(RichTextBoxDataObjectValue) Then
    AllowedEffects = vbDropEffectNone
    If RichTextBoxHandle <> 0 Then
        Dim OLEInstance As OLEGuids.IRichEditOle
        Set OLEInstance = Me.GetOLEInterface
        If Not OLEInstance Is Nothing Then
            If IsEmpty(RichTextBoxDataObjectFormat) Then
                Data.SetData RichTextBoxDataObjectValue
            Else
                Data.SetData RichTextBoxDataObjectValue, RichTextBoxDataObjectFormat
            End If
            Dim PropDataObject As OLEGuids.IDataObject, TempObj As Object
            ' VB stores a pointer to the IDataObject interface.
            ' OLEStartDrag event: In the 5th Long from the object pointer. (+20)
            ' OLEDragDrop event: In the 4th Long from the object pointer. (+16)
            CopyMemory TempObj, ByVal UnsignedAdd(ObjPtr(Data), 20), 4
            Set PropDataObject = TempObj
            CopyMemory TempObj, 0&, 4
            If Not PropDataObject Is Nothing Then
                Dim PropOleObject As OLEGuids.IOleObject, PropClientSite As OLEGuids.IOleClientSite, PropStorage As OLEGuids.IStorage
                Set PropClientSite = OLEInstance.GetClientSite
                StgCreateDocFile 0, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE Or STGM_DELETEONRELEASE, 0, PropStorage
                Const IID_IOleObject As String = "{00000112-0000-0000-C000-000000000046}"
                Dim IID As OLEGuids.OLECLSID
                CLSIDFromString StrPtr(IID_IOleObject), IID
                Select Case OleQueryCreateFromData(PropDataObject)
                    Case OLE_S_STATIC
                        OleCreateStaticFromData PropDataObject, IID, OLERENDER_DRAW, 0, PropClientSite, PropStorage, PropOleObject
                    Case S_OK
                        OleCreateFromData PropDataObject, IID, OLERENDER_DRAW, 0, PropClientSite, PropStorage, PropOleObject
                End Select
                If Not PropOleObject Is Nothing Then
                    OleSetContainedObject PropOleObject, 1
                    Dim REOBJ As REOBJECT
                    With REOBJ
                    .cbStruct = LenB(REOBJ)
                    LSet .riid = IID
                    .dvAspect = DVASPECT_CONTENT
                    .CharPos = REO_CP_SELECTION
                    .dwFlags = REO_DYNAMICSIZE Or REO_RESIZABLE Or REO_BELOWBASELINE
                    .Size.CX = 0
                    .Size.CY = 0
                    .dwUser = 0
                    Set .pStorage = PropStorage
                    Set .pOleSite = PropClientSite
                    Set .pOleObject = PropOleObject
                    End With
                    OLEInstance.InsertObject REOBJ
                End If
            End If
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
If RichTextBoxHandle <> 0 Then MoveWindow RichTextBoxHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfaceControl)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyRichTextBox
Call ComCtlsReleaseShellMod
Call RtfReleaseRichedMod
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
hWnd = RichTextBoxHandle
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
OldFontHandle = RichTextBoxFontHandle
RichTextBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, WM_SETFONT, RichTextBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = RichTextBoxFontHandle
RichTextBoxFontHandle = CreateGDIFontFromOLEFont(PropFont)
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, WM_SETFONT, RichTextBoxFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
RichTextBoxEnabledVisualStyles = EnabledVisualStyles()
If RichTextBoxHandle <> 0 And RichTextBoxEnabledVisualStyles = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles RichTextBoxHandle
    Else
        RemoveVisualStyles RichTextBoxHandle
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
If RichTextBoxHandle <> 0 Then EnableWindow RichTextBoxHandle, IIf(Value = True, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDragDrop() As Boolean
Attribute OLEDragDrop.VB_Description = "Returns/Sets whether this object can act as an OLE drag source and drop target."
OLEDragDrop = PropOLEDragDrop
End Property

Public Property Let OLEDragDrop(ByVal Value As Boolean)
PropOLEDragDrop = Value
If RichTextBoxHandle <> 0 Then Call ReCreateRichTextBox
UserControl.PropertyChanged "OLEDragDrop"
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
If RichTextBoxDesignMode = False Then Call RefreshMousePointer
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
        If RichTextBoxDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If RichTextBoxDesignMode = False Then Call RefreshMousePointer
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
If RichTextBoxHandle <> 0 Then Call ReCreateRichTextBox(NoStreamStringOutIn:=CBool(Me.TextMode = RtfTextModeRichText))
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
If RichTextBoxHandle <> 0 Then
    Dim dwStyle As Long, dwExStyle As Long
    dwStyle = GetWindowLong(RichTextBoxHandle, GWL_STYLE)
    dwExStyle = GetWindowLong(RichTextBoxHandle, GWL_EXSTYLE)
    If PropBorderStyle = vbFixedSingle Then
        If Not (dwStyle And ES_SUNKEN) = ES_SUNKEN Then dwStyle = dwStyle Or ES_SUNKEN
        If Not (dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
    Else
        If (dwStyle And ES_SUNKEN) = ES_SUNKEN Then dwStyle = dwStyle And Not ES_SUNKEN
        If (dwExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
    End If
    SetWindowLong RichTextBoxHandle, GWL_STYLE, dwStyle
    SetWindowLong RichTextBoxHandle, GWL_EXSTYLE, dwExStyle
    Call ComCtlsFrameChanged(RichTextBoxHandle)
End If
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object. Only applicable if the enabled property is set to true."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If RichTextBoxHandle <> 0 Then
    SendMessage RichTextBoxHandle, EM_SETBKGNDCOLOR, 0, ByVal WinColor(PropBackColor)
    
    #If ImplementThemedBorder = True Then
    
    ' Redraw the border to consider the new back color for the themed border, if any.
    RedrawWindow RichTextBoxHandle, 0, 0, RDW_FRAME Or RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_NOCHILDREN
    
    #End If
    
End If
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents can be edited."
If RichTextBoxHandle <> 0 Then
    Locked = CBool((GetWindowLong(RichTextBoxHandle, GWL_STYLE) And ES_READONLY) <> 0)
Else
    Locked = PropLocked
End If
End Property

Public Property Let Locked(ByVal Value As Boolean)
PropLocked = Value
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_SETREADONLY, IIf(PropLocked = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "Locked"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value indicating if the selection in an edit control is hidden when the control loses focus."
HideSelection = PropHideSelection
End Property

Public Property Let HideSelection(ByVal Value As Boolean)
PropHideSelection = Value
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_HIDESELECTION, IIf(PropHideSelection = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "HideSelection"
End Property

Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
If RichTextBoxHandle <> 0 Then
    PasswordChar = ChrW(SendMessage(RichTextBoxHandle, EM_GETPASSWORDCHAR, 0, ByVal 0&))
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
    If RichTextBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If RichTextBoxHandle <> 0 Then
    SendMessage RichTextBoxHandle, EM_SETPASSWORDCHAR, PropPasswordChar, ByVal 0&
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
If RichTextBoxHandle <> 0 Then Call ReCreateRichTextBox
UserControl.PropertyChanged "UseSystemPasswordChar"
End Property

Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
MultiLine = PropMultiLine
End Property

Public Property Let MultiLine(ByVal Value As Boolean)
If RichTextBoxDesignMode = False Then
    Err.Raise Number:=382, Description:="MultiLine property is read-only at run time"
Else
    PropMultiLine = Value
    If RichTextBoxHandle <> 0 Then Call ReCreateRichTextBox
End If
UserControl.PropertyChanged "MultiLine"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
MaxLength = PropMaxLength
End Property

Public Property Let MaxLength(ByVal Value As Long)
If Value < 0 Then
    If RichTextBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropMaxLength = Value
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_EXLIMITTEXT, 0, ByVal PropMaxLength
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
        If RichTextBoxHandle <> 0 Then Call ReCreateRichTextBox
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "ScrollBars"
End Property

Public Property Get WantReturn() As Boolean
Attribute WantReturn.VB_Description = "Returns/sets a value that determines when the user presses RETURN to perform the default button or to advance to the next line. This property applies only to a multiline rich text box and when there is any default button on the form."
WantReturn = PropWantReturn
End Property

Public Property Let WantReturn(ByVal Value As Boolean)
If PropWantReturn = Value Then Exit Property
PropWantReturn = Value
If RichTextBoxHandle <> 0 And RichTextBoxDesignMode = False Then
    ' It is not possible (in VB6) to achieve this when specifying ES_WANTRETURN.
    Dim PropOleObject As OLEGuids.IOleObject
    Dim PropClientSite As OLEGuids.IOleClientSite
    Dim PropUnknown As IUnknown
    Dim PropControlSite As OLEGuids.IOleControlSite
    On Error Resume Next
    Set PropOleObject = Me
    Set PropClientSite = PropOleObject.GetClientSite
    Set PropUnknown = PropClientSite
    Set PropControlSite = PropUnknown
    PropControlSite.OnControlInfoChanged
    If GetFocus() = RichTextBoxHandle Then
        ' If focus is on the control then force the change immediately.
        PropControlSite.OnFocus 1
    End If
    On Error GoTo 0
End If
UserControl.PropertyChanged "WantReturn"
End Property

Public Property Get DisableNoScroll() As Boolean
Attribute DisableNoScroll.VB_Description = "Returns/sets a value that determines whether scroll bars are disabled instead of hided when they are not needed."
DisableNoScroll = PropDisableNoScroll
End Property

Public Property Let DisableNoScroll(ByVal Value As Boolean)
If RichTextBoxDesignMode = False Then
    Err.Raise Number:=382, Description:="DisableNoScroll property is read-only at run time"
Else
    PropDisableNoScroll = Value
    If RichTextBoxHandle <> 0 Then Call ReCreateRichTextBox
End If
UserControl.PropertyChanged "DisableNoScroll"
End Property

Public Property Get AutoURLDetect() As Boolean
Attribute AutoURLDetect.VB_Description = "Returns/sets a value indicating if automatic detection of hyperlinks is enabled or disabled."
If RichTextBoxHandle <> 0 Then
    AutoURLDetect = CBool(SendMessage(RichTextBoxHandle, EM_GETAUTOURLDETECT, 0, ByVal 0&) = 1)
Else
    AutoURLDetect = PropAutoURLDetect
End If
End Property

Public Property Let AutoURLDetect(ByVal Value As Boolean)
PropAutoURLDetect = Value
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_AUTOURLDETECT, IIf(PropAutoURLDetect = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "AutoURLDetect"
End Property

Public Property Get BulletIndent() As Single
Attribute BulletIndent.VB_Description = "Returns/sets the amount of indent used when a paragraph has the bullet style."
BulletIndent = UserControl.ScaleX(PropBulletIndent, vbPixels, vbContainerSize)
End Property

Public Property Let BulletIndent(ByVal Value As Single)
If Value < 0 Then
    If RichTextBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim LngValue As Long, ErrValue As Long
On Error Resume Next
LngValue = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
ErrValue = Err.Number
On Error GoTo 0
If LngValue >= 0 And ErrValue = 0 Then
    PropBulletIndent = LngValue
Else
    If RichTextBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
UserControl.PropertyChanged "BulletIndent"
End Property

Public Property Get SelectionBar() As Boolean
Attribute SelectionBar.VB_Description = "Returns/sets a value that determines whether or not the control adds space to the left margin where the cursor changes to a right-up arrow, allowing the user to select full lines of text."
SelectionBar = PropSelectionBar
End Property

Public Property Let SelectionBar(ByVal Value As Boolean)
PropSelectionBar = Value
If RichTextBoxHandle <> 0 Then
    Dim Flags As Long
    Flags = SendMessage(RichTextBoxHandle, EM_GETOPTIONS, 0, ByVal 0&)
    If PropSelectionBar = True Then
        If Not (Flags And ECO_SELECTIONBAR) = ECO_SELECTIONBAR Then Flags = Flags Or ECO_SELECTIONBAR
    Else
        If (Flags And ECO_SELECTIONBAR) = ECO_SELECTIONBAR Then Flags = Flags And Not ECO_SELECTIONBAR
    End If
    SendMessage RichTextBoxHandle, EM_SETOPTIONS, ECOOP_SET, ByVal Flags
End If
UserControl.PropertyChanged "SelectionBar"
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the file name of the file loaded into the rich text box control at design time."
Attribute FileName.VB_ProcData.VB_Invoke_Property = "PPRichTextBoxGeneral"
FileName = PropFileName
End Property

Public Property Let FileName(ByVal Value As String)
If Value = vbNullString Then
    PropFileName = vbNullString
    If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, WM_SETTEXT, 0, ByVal 0&
Else
    If FileExists(Value) = True Then
        PropFileName = Value
        Dim hFile As Long, Length As Long
        Dim B1(0 To 1) As Byte, B2(0 To 2) As Byte
        hFile = CreateFile(StrPtr("\\?\" & IIf(Left$(PropFileName, 2) = "\\", "UNC\" & Mid$(PropFileName, 3), PropFileName)), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
        If hFile <> INVALID_HANDLE_VALUE Then
            Length = GetFileSize(hFile, 0) ' File size >= 2^31 not supported.
            If Length >= 2 Then
                ReadFile hFile, VarPtr(B1(0)), 2, 0, 0
                If Length >= 5 Then ReadFile hFile, VarPtr(B2(0)), 3, 0, 0
            End If
            CloseHandle hFile
        End If
        If B1(0) = &HFF And B1(1) = &HFE Then ' UTF-16 BOM
            Me.LoadFile PropFileName, RtfLoadSaveFormatUnicodeText
        Else
            If B1(0) = &H7B And B1(1) = &H5C And StrComp(StrConv(B2(), vbUnicode), "rtf", vbTextCompare) = 0 Then
                If Me.TextMode = RtfTextModeRichText Then
                    Me.LoadFile PropFileName, RtfLoadSaveFormatRTF
                Else
                    PropFileName = vbNullString
                    Exit Property
                End If
            Else
                Me.LoadFile PropFileName, RtfLoadSaveFormatText
            End If
        End If
    Else
        If RichTextBoxDesignMode = True Then
            MsgBox "The specified file name cannot be accessed or is invalid.", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise Number:=75, Description:="The specified file name cannot be accessed or is invalid"
        End If
    End If
End If
UserControl.PropertyChanged "FileName"
End Property

Public Property Get TextMode() As RtfTextModeConstants
Attribute TextMode.VB_Description = "Returns/sets the text mode."
If RichTextBoxHandle <> 0 Then
    If (SendMessage(RichTextBoxHandle, EM_GETTEXTMODE, 0, ByVal 0&) And TM_RICHTEXT) <> 0 Then
        TextMode = RtfTextModeRichText
    Else
        TextMode = RtfTextModePlainText
    End If
Else
    TextMode = PropTextMode
End If
End Property

Public Property Let TextMode(ByVal Value As RtfTextModeConstants)
Select Case Value
    Case RtfTextModeRichText, RtfTextModePlainText
        If RichTextBoxDesignMode = True Then PropFileName = vbNullString
        PropTextMode = Value
    Case Else
        Err.Raise 380
End Select
If RichTextBoxHandle <> 0 Then
    SendMessage RichTextBoxHandle, WM_SETTEXT, 0, ByVal 0&
    SendMessage RichTextBoxHandle, EM_SETTEXTMODE, IIf(PropTextMode = RtfTextModeRichText, TM_RICHTEXT, TM_PLAINTEXT), ByVal 0&
End If
UserControl.PropertyChanged "TextMode"
End Property

Public Property Get UndoLimit() As Long
Attribute UndoLimit.VB_Description = "Returns/sets the maximum number of actions that can be stored in the undo queue. A value of 0 indicates that the undo feature is disabled."
UndoLimit = PropUndoLimit
End Property

Public Property Let UndoLimit(ByVal Value As Long)
If Value < 0 Then
    If RichTextBoxDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If RichTextBoxHandle <> 0 Then
    If Value < 1 Then
        SendMessage RichTextBoxHandle, EM_SETTEXTMODE, TM_SINGLELEVELUNDO, ByVal 0&
        PropUndoLimit = SendMessage(RichTextBoxHandle, EM_SETUNDOLIMIT, 0, ByVal 0&)
    ElseIf Value = 1 Then
        SendMessage RichTextBoxHandle, EM_SETTEXTMODE, TM_SINGLELEVELUNDO, ByVal 0&
        PropUndoLimit = SendMessage(RichTextBoxHandle, EM_SETUNDOLIMIT, Value, ByVal 0&)
    Else
        SendMessage RichTextBoxHandle, EM_SETTEXTMODE, TM_MULTILEVELUNDO, ByVal 0&
        PropUndoLimit = SendMessage(RichTextBoxHandle, EM_SETUNDOLIMIT, Value, ByVal 0&)
    End If
End If
UserControl.PropertyChanged "UndoLimit"
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
If RichTextBoxHandle <> 0 And RichTextBoxDesignMode = False Then
    If GetFocus() = RichTextBoxHandle Then Call ComCtlsSetIMEMode(RichTextBoxHandle, RichTextBoxIMCHandle, PropIMEMode)
End If
UserControl.PropertyChanged "IMEMode"
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
If RichTextBoxHandle <> 0 And RichTextBoxDesignMode = False Then
    SendMessage RichTextBoxHandle, WM_KEYDOWN, vbKeyInsert, ByVal 0&
    SendMessage RichTextBoxHandle, WM_KEYUP, vbKeyInsert, ByVal 0&
Else
    If PropAllowOverType = True Then PropOverTypeMode = Value Else PropOverTypeMode = False
End If
UserControl.PropertyChanged "OverTypeMode"
End Property

Public Property Get UseCrLf() As Boolean
Attribute UseCrLf.VB_Description = "Returns/sets a value that determines whether or not the control translates each Cr into a CrLf for the text property."
UseCrLf = PropUseCrLf
End Property

Public Property Let UseCrLf(ByVal Value As Boolean)
PropUseCrLf = Value
UserControl.PropertyChanged "UseCrLf"
End Property

Private Sub CreateRichTextBox()
If RichTextBoxHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropOLEDragDrop = False Then dwStyle = dwStyle Or ES_NOOLEDRAGDROP
If PropRightToLeft = True Then dwExStyle = dwExStyle Or WS_EX_RTLREADING Or WS_EX_RIGHT Or WS_EX_LEFTSCROLLBAR
If PropBorderStyle = vbFixedSingle Then
    dwStyle = dwStyle Or ES_SUNKEN
    dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
End If
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
If PropDisableNoScroll = True Then dwStyle = dwStyle Or ES_DISABLENOSCROLL
If PropSelectionBar = True Then dwStyle = dwStyle Or ES_SELECTIONBAR
Dim ClassName As String
ClassName = RtfGetClassName()
RichTextBoxHandle = CreateWindowEx(dwExStyle, StrPtr(ClassName), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If RichTextBoxHandle <> 0 Then
    If PropPasswordChar <> 0 And PropUseSystemPasswordChar = False Then SendMessage RichTextBoxHandle, EM_SETPASSWORDCHAR, PropPasswordChar, ByVal 0&
    SendMessage RichTextBoxHandle, EM_EXLIMITTEXT, 0, ByVal PropMaxLength
    SendMessage RichTextBoxHandle, EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY, ByVal TO_ADVANCEDTYPOGRAPHY
    If PropTextMode = RtfTextModePlainText Then SendMessage RichTextBoxHandle, EM_SETTEXTMODE, TM_PLAINTEXT, ByVal 0&
    Dim This As OLEGuids.IRichEditOleCallback
    Set This = RtfOleCallback(Me)
    If Not This Is Nothing Then
        RichTextBoxIsOleCallback = CBool(SendMessage(RichTextBoxHandle, EM_SETOLECALLBACK, 0, ByVal ObjPtr(This)) <> 0)
    Else
        RichTextBoxIsOleCallback = False
    End If
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.BackColor = PropBackColor
Me.AutoURLDetect = PropAutoURLDetect
If PropUndoLimit <> 100 Then Me.UndoLimit = PropUndoLimit
If RichTextBoxDesignMode = False Then
    If RichTextBoxHandle <> 0 Then
        SendMessage RichTextBoxHandle, EM_SETEVENTMASK, 0, ByVal ENM_CHANGE Or ENM_SCROLL Or ENM_SELCHANGE Or ENM_DRAGDROPDONE Or ENM_LINK Or ENM_PROTECTED
        SendMessage RichTextBoxHandle, EM_SETEDITSTYLE, SES_BEEPONMAXTEXT, ByVal SES_BEEPONMAXTEXT
        If PropAllowOverType = True And PropOverTypeMode = True Then
            SendMessage RichTextBoxHandle, WM_KEYDOWN, vbKeyInsert, ByVal 0&
            SendMessage RichTextBoxHandle, WM_KEYUP, vbKeyInsert, ByVal 0&
        End If
        Call ComCtlsSetSubclass(RichTextBoxHandle, Me, 1)
    End If
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
    If RichTextBoxHandle <> 0 Then Call ComCtlsCreateIMC(RichTextBoxHandle, RichTextBoxIMCHandle)
End If
End Sub

Private Sub ReCreateRichTextBox(Optional ByVal NoStreamStringOutIn As Boolean)
Dim Buffer As String, Flags As Long
If Me.TextMode = RtfTextModeRichText Then Flags = SF_RTF Else Flags = SF_TEXT Or SF_UNICODE
If RichTextBoxDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Dim RECR As RECHARRANGE, P As POINTAPI
    If RichTextBoxHandle <> 0 Then
        SendMessage RichTextBoxHandle, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
        SendMessage RichTextBoxHandle, EM_GETSCROLLPOS, 0, ByVal VarPtr(P)
        If PropScrollBars = vbVertical Or PropScrollBars = vbSBNone Then P.X = 0
        If PropScrollBars = vbHorizontal Or PropScrollBars = vbSBNone Then P.Y = 0
        If NoStreamStringOutIn = False Then StreamStringOut Buffer, Flags
    End If
    Call DestroyRichTextBox
    Call CreateRichTextBox
    Call UserControl_Resize
    If RichTextBoxHandle <> 0 Then
        If NoStreamStringOutIn = False Then StreamStringIn Buffer, Flags
        SendMessage RichTextBoxHandle, EM_EXSETSEL, 0, ByVal VarPtr(RECR)
        If P.X > 0 Or P.Y > 0 Then SendMessage RichTextBoxHandle, EM_SETSCROLLPOS, 0, ByVal VarPtr(P)
    End If
    If Locked = True Then LockWindowUpdate 0
    Me.Refresh
Else
    If NoStreamStringOutIn = False Then StreamStringOut Buffer, Flags
    Call DestroyRichTextBox
    Call CreateRichTextBox
    Call UserControl_Resize
    If PropFileName = vbNullString Then
        If NoStreamStringOutIn = False Then StreamStringIn Buffer, Flags
    Else
        Me.FileName = PropFileName
    End If
End If
End Sub

Private Sub DestroyRichTextBox()
If RichTextBoxHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(RichTextBoxHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
Call ComCtlsDestroyIMC(RichTextBoxHandle, RichTextBoxIMCHandle)
If RichTextBoxIsOleCallback = True Then RichTextBoxIsOleCallback = Not CBool(SendMessage(RichTextBoxHandle, EM_SETOLECALLBACK, 0, ByVal 0&) <> 0)
ShowWindow RichTextBoxHandle, SW_HIDE
SetParent RichTextBoxHandle, 0
DestroyWindow RichTextBoxHandle
RichTextBoxHandle = 0
If RichTextBoxFontHandle <> 0 Then
    DeleteObject RichTextBoxFontHandle
    RichTextBoxFontHandle = 0
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
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, WM_COPY, 0, ByVal 0&
End Sub

Public Sub Cut()
Attribute Cut.VB_Description = "Method to delete (cut) the current selection and copy the deleted text to the clipboard."
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, WM_CUT, 0, ByVal 0&
End Sub

Public Sub Paste()
Attribute Paste.VB_Description = "Method to copy the current content of the clipboard at the current caret position."
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, WM_PASTE, 0, ByVal 0&
End Sub

Public Function CanPaste() As Boolean
Attribute CanPaste.VB_Description = "Determines whether there is any format currently on the clipboard that can be pasted."
If RichTextBoxHandle <> 0 Then CanPaste = CBool(SendMessage(RichTextBoxHandle, EM_CANPASTE, 0, ByVal 0&) <> 0)
End Function

Public Sub PasteSpecial(ByVal wFormat As Long)
Attribute PasteSpecial.VB_Description = "Pastes a specific clipboard format in a rich text box control."
If RichTextBoxHandle <> 0 Then
    If wFormat = vbCFRTF Then wFormat = RegisterClipboardFormat(StrPtr("Rich Text Format"))
    SendMessage RichTextBoxHandle, EM_PASTESPECIAL, wFormat, ByVal 0&
End If
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Method to delete (clear) the current selection."
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, WM_CLEAR, 0, ByVal 0&
End Sub

Public Sub Undo()
Attribute Undo.VB_Description = "Undoes the last operation, if any."
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_UNDO, 0, ByVal 0&
End Sub

Public Property Get UndoType() As RtfActionTypeConstants
Attribute UndoType.VB_Description = "Retrieves the type of the next undo action, if any."
If RichTextBoxHandle <> 0 Then UndoType = SendMessage(RichTextBoxHandle, EM_GETUNDONAME, 0, ByVal 0&)
End Property

Public Function CanUndo() As Boolean
Attribute CanUndo.VB_Description = "Determines whether there are any actions in the undo queue."
If RichTextBoxHandle <> 0 Then CanUndo = CBool(SendMessage(RichTextBoxHandle, EM_CANUNDO, 0, ByVal 0&) <> 0)
End Function

Public Sub StopUndoAction()
Attribute StopUndoAction.VB_Description = "Stops the rich text box control from collecting additional typing actions into the current undo action."
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_STOPGROUPTYPING, 0, ByVal 0&
End Sub

Public Sub ResetUndoQueue()
Attribute ResetUndoQueue.VB_Description = "Resets the undo queue."
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
End Sub

Public Sub Redo()
Attribute Redo.VB_Description = "Redoes the next action in the redo queue, if any."
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_REDO, 0, ByVal 0&
End Sub

Public Property Get RedoType() As RtfActionTypeConstants
Attribute RedoType.VB_Description = "Retrieves the type of the next redo action, if any."
If RichTextBoxHandle <> 0 Then RedoType = SendMessage(RichTextBoxHandle, EM_GETREDONAME, 0, ByVal 0&)
End Property

Public Function CanRedo() As Boolean
Attribute CanRedo.VB_Description = "Determines whether there are any actions in the redo queue."
If RichTextBoxHandle <> 0 Then CanRedo = CBool(SendMessage(RichTextBoxHandle, EM_CANREDO, 0, ByVal 0&) <> 0)
End Function

Public Property Get Modified() As Boolean
Attribute Modified.VB_Description = "Setting the text property will reset this property to false. Any typing in the control will set the property to true."
Attribute Modified.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then Modified = CBool(SendMessage(RichTextBoxHandle, EM_GETMODIFY, 0, ByVal 0&) <> 0)
End Property

Public Property Let Modified(ByVal Value As Boolean)
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_SETMODIFY, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Function GetTextRange(ByVal Min As Long, ByVal Max As Long) As String
If Min > Max Then Err.Raise 380
If RichTextBoxHandle <> 0 Then
    Dim RETR As RETEXTRANGE, Buffer As String, Length As Long
    RETR.CharRange.Min = Min
    RETR.CharRange.Max = Max
    Buffer = String$(RETR.CharRange.Max - RETR.CharRange.Min + 1, vbNullChar)
    RETR.lpstrText = StrPtr(Buffer)
    Length = SendMessage(RichTextBoxHandle, EM_GETTEXTRANGE, 0, ByVal VarPtr(RETR))
    If Length > 0 Then GetTextRange = Left$(Buffer, Length)
End If
End Function

Public Function Find(ByVal Text As String, Optional ByVal Min As Long, Optional ByVal Max As Long = -1, Optional ByVal Options As RtfFindOptionConstants) As Long
Attribute Find.VB_Description = "Finds text within a rich text box control."
If RichTextBoxHandle <> 0 Then
    Dim REFTEX As REFINDTEXTEX, dwOptions As Long
    With REFTEX
    With .CharRange
    If Min >= 0 Then
        .Min = Min
    Else
        Err.Raise 380
    End If
    If Max >= -1 Then
        .Max = Max
    Else
        Err.Raise 380
    End If
    End With
    .lpstrText = StrPtr(Text)
    Const FR_DOWN As Long = &H1
    dwOptions = FR_DOWN
    If (Options And RtfFindOptionWholeWord) <> 0 Then dwOptions = dwOptions Or FR_WHOLEWORD
    If (Options And RtfFindOptionMatchCase) <> 0 Then dwOptions = dwOptions Or FR_MATCHCASE
    If (Options And RtfFindOptionReverse) <> 0 Then dwOptions = dwOptions And Not FR_DOWN
    Find = SendMessage(RichTextBoxHandle, EM_FINDTEXTEX, dwOptions, ByVal VarPtr(REFTEX))
    If (Options And RtfFindOptionNoHighlight) = 0 And Find <> -1 Then SendMessage RichTextBoxHandle, EM_EXSETSEL, 0, ByVal VarPtr(.CharRangeText)
    End With
End If
End Function

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_ProcData.VB_Invoke_Property = "PPRichTextBoxText"
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "121c"
If RichTextBoxHandle <> 0 Then
    Dim REGTLEX As REGETTEXTLENGTHEX, Length As Long
    REGTLEX.Flags = GTL_PRECISE Or GTL_NUMCHARS
    If PropUseCrLf = True Then REGTLEX.Flags = REGTLEX.Flags Or GTL_USECRLF
    REGTLEX.CodePage = CP_UNICODE
    Length = SendMessage(RichTextBoxHandle, EM_GETTEXTLENGTHEX, VarPtr(REGTLEX), ByVal 0&)
    If Length > 0 Then
        Dim REGTEX As REGETTEXTEX, Buffer As String
        REGTEX.cbSize = (Length + 1) * 2
        If PropUseCrLf = False Then REGTEX.Flags = GT_DEFAULT Else REGTEX.Flags = GT_USECRLF
        REGTEX.CodePage = CP_UNICODE
        Buffer = String$(Length, vbNullChar)
        Length = SendMessage(RichTextBoxHandle, EM_GETTEXTEX, VarPtr(REGTEX), ByVal StrPtr(Buffer))
        If Length > 0 Then Text = Left$(Buffer, Length)
    End If
End If
End Property

Public Property Let Text(ByVal Value As String)
If RichTextBoxDesignMode = True Then PropFileName = vbNullString
If RichTextBoxHandle <> 0 Then
    Dim RESTEX As RESETTEXTEX
    RESTEX.Flags = ST_UNICODE
    RESTEX.CodePage = CP_UNICODE
    SendMessage RichTextBoxHandle, EM_SETTEXTEX, VarPtr(RESTEX), ByVal StrPtr(Value)
End If
UserControl.PropertyChanged "Text"
End Property

Public Property Get TextLength() As Long
Attribute TextLength.VB_Description = "Returns the length of the text."
Attribute TextLength.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim REGTLEX As REGETTEXTLENGTHEX
    REGTLEX.Flags = GTL_PRECISE Or GTL_NUMCHARS
    If PropUseCrLf = True Then REGTLEX.Flags = REGTLEX.Flags Or GTL_USECRLF
    REGTLEX.CodePage = CP_UNICODE
    TextLength = SendMessage(RichTextBoxHandle, EM_GETTEXTLENGTHEX, VarPtr(REGTLEX), ByVal 0&)
End If
End Property

Public Property Get TextRTF() As String
Attribute TextRTF.VB_Description = "Returns/sets the text of a rich text box control, including all .RTF code."
Attribute TextRTF.VB_MemberFlags = "143c"
StreamStringOut TextRTF, SF_RTF
End Property

Public Property Let TextRTF(ByVal Value As String)
StreamStringIn Value, SF_RTF
UserControl.PropertyChanged "TextRTF"
End Property

Public Property Get Default() As String
Attribute Default.VB_UserMemId = 0
Attribute Default.VB_MemberFlags = "40"
Default = Me.TextRTF
End Property

Public Property Let Default(ByVal Value As String)
Me.TextRTF = Value
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECR As RECHARRANGE, Buffer As String, Length As Long
    SendMessage RichTextBoxHandle, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
    Buffer = String$(RECR.Max - RECR.Min + 1, vbNullChar)
    Length = SendMessage(RichTextBoxHandle, EM_GETSELTEXT, 0, ByVal StrPtr(Buffer))
    If Length > 0 Then SelText = Left$(Buffer, Length)
End If
End Property

Public Property Let SelText(ByVal Value As String)
If RichTextBoxHandle <> 0 Then
    If StrPtr(Value) = 0 Then Value = ""
    SendMessage RichTextBoxHandle, EM_REPLACESEL, 1, ByVal StrPtr(Value)
End If
End Property

Public Property Get SelRTF() As String
Attribute SelRTF.VB_Description = "Returns/sets the string containing the currently selected text, including all .RTF code."
Attribute SelRTF.VB_MemberFlags = "400"
StreamStringOut SelRTF, SF_RTF Or SFF_SELECTION
End Property

Public Property Let SelRTF(ByVal Value As String)
StreamStringIn Value, SF_RTF Or SFF_SELECTION
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected; indicates the position of the insertion point if no text is selected."
Attribute SelStart.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECR As RECHARRANGE
    SendMessage RichTextBoxHandle, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
    SelStart = RECR.Min
End If
End Property

Public Property Let SelStart(ByVal Value As Long)
If RichTextBoxHandle <> 0 Then
    If Value >= 0 Then
        Dim RECR As RECHARRANGE
        RECR.Min = Value
        RECR.Max = Value
        SendMessage RichTextBoxHandle, EM_EXSETSEL, 0, ByVal VarPtr(RECR)
        Me.ScrollToCaret
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECR As RECHARRANGE
    SendMessage RichTextBoxHandle, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
    SelLength = RECR.Max - RECR.Min
End If
End Property

Public Property Let SelLength(ByVal Value As Long)
If RichTextBoxHandle <> 0 Then
    If Value >= 0 Then
        Dim RECR As RECHARRANGE
        SendMessage RichTextBoxHandle, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
        RECR.Max = RECR.Min + Value
        SendMessage RichTextBoxHandle, EM_EXSETSEL, 0, ByVal VarPtr(RECR)
        Me.ScrollToCaret
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelAlignment() As Variant
Attribute SelAlignment.VB_Description = "Returns/sets a value that controls the alignment of a paragraph."
Attribute SelAlignment.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_ALIGNMENT
    If (SendMessage(RichTextBoxHandle, EM_GETPARAFORMAT, 0, ByVal VarPtr(REPF2)) And PFM_ALIGNMENT) <> 0 Then
        Select Case .Alignment
            Case PFA_LEFT
                SelAlignment = RtfSelAlignmentLeft
            Case PFA_RIGHT
                SelAlignment = RtfSelAlignmentRight
            Case PFA_CENTER
                SelAlignment = RtfSelAlignmentCenter
            Case PFA_JUSTIFY
                SelAlignment = RtfSelAlignmentJustified
        End Select
    Else
        SelAlignment = Null
    End If
    End With
End If
End Property

Public Property Let SelAlignment(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_ALIGNMENT
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            Select Case Value
                Case RtfSelAlignmentLeft
                    .Alignment = PFA_LEFT
                Case RtfSelAlignmentRight
                    .Alignment = PFA_RIGHT
                Case RtfSelAlignmentCenter
                    .Alignment = PFA_CENTER
                Case RtfSelAlignmentJustified
                    .Alignment = PFA_JUSTIFY
                Case Else
                    Err.Raise 380
            End Select
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETPARAFORMAT, 0, ByVal VarPtr(REPF2)
End If
End Property

Public Property Get SelBold() As Variant
Attribute SelBold.VB_Description = "Returns/sets the bold format of the currently selected text."
Attribute SelBold.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_BOLD
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_BOLD) <> 0 Then
        SelBold = CBool((.dwEffects And CFE_BOLD) = CFE_BOLD)
    Else
        SelBold = Null
    End If
    End With
End If
End Property

Public Property Let SelBold(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_BOLD
    Select Case VarType(Value)
        Case vbBoolean
            If Value = True Then .dwEffects = CFE_BOLD Else .dwEffects = 0
        Case vbNull
            .dwEffects = CFE_BOLD
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            If CBool(Value) = True Then .dwEffects = CFE_BOLD Else .dwEffects = 0
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelItalic() As Variant
Attribute SelItalic.VB_Description = "Returns/set the italic format of the currently selected text."
Attribute SelItalic.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_ITALIC
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_ITALIC) <> 0 Then
        SelItalic = CBool((.dwEffects And CFE_ITALIC) = CFE_ITALIC)
    Else
        SelItalic = Null
    End If
    End With
End If
End Property

Public Property Let SelItalic(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_ITALIC
    Select Case VarType(Value)
        Case vbBoolean
            If Value = True Then .dwEffects = CFE_ITALIC Else .dwEffects = 0
        Case vbNull
            .dwEffects = CFE_ITALIC
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            If CBool(Value) = True Then .dwEffects = CFE_ITALIC Else .dwEffects = 0
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelStrikethru() As Variant
Attribute SelStrikethru.VB_Description = "Returns/set the strikethru format of the currently selected text."
Attribute SelStrikethru.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_STRIKEOUT
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_STRIKEOUT) <> 0 Then
        SelStrikethru = CBool((.dwEffects And CFE_STRIKEOUT) = CFE_STRIKEOUT)
    Else
        SelStrikethru = Null
    End If
    End With
End If
End Property

Public Property Let SelStrikethru(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_STRIKEOUT
    Select Case VarType(Value)
        Case vbBoolean
            If Value = True Then .dwEffects = CFE_STRIKEOUT Else .dwEffects = 0
        Case vbNull
            .dwEffects = CFE_STRIKEOUT
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            If CBool(Value) = True Then .dwEffects = CFE_STRIKEOUT Else .dwEffects = 0
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelUnderline() As Variant
Attribute SelUnderline.VB_Description = "Returns/set the underline format of the currently selected text."
Attribute SelUnderline.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_UNDERLINE
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_UNDERLINE) <> 0 Then
        SelUnderline = CBool((.dwEffects And CFE_UNDERLINE) = CFE_UNDERLINE)
    Else
        SelUnderline = Null
    End If
    End With
End If
End Property

Public Property Let SelUnderline(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_UNDERLINE
    Select Case VarType(Value)
        Case vbBoolean
            If Value = True Then .dwEffects = CFE_UNDERLINE Else .dwEffects = 0
        Case vbNull
            .dwEffects = CFE_UNDERLINE
        Case vbLong, vbInteger, vbByte, vbSingle, vbDouble
            If CBool(Value) = True Then .dwEffects = CFE_UNDERLINE Else .dwEffects = 0
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelBullet() As Variant
Attribute SelBullet.VB_Description = "Returns/sets a value that determines if a paragraph in the current selection or insertion point has the bullet style."
Attribute SelBullet.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_NUMBERING
    If (SendMessage(RichTextBoxHandle, EM_GETPARAFORMAT, 0, ByVal VarPtr(REPF2)) And PFM_NUMBERING) <> 0 Then
        SelBullet = CBool((.Numbering And PFN_BULLET) = PFN_BULLET)
    Else
        SelBullet = Null
    End If
    End With
End If
End Property

Public Property Let SelBullet(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_NUMBERING Or PFM_OFFSET
    Select Case VarType(Value)
        Case vbBoolean
            If Value = True Then .Numbering = PFN_BULLET Else .Numbering = 0
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            If CBool(Value) = True Then .Numbering = PFN_BULLET Else .Numbering = 0
        Case Else
            Err.Raise 13
    End Select
    .DXOffset = UserControl.ScaleX(PropBulletIndent, vbPixels, vbTwips)
    End With
    SendMessage RichTextBoxHandle, EM_SETPARAFORMAT, 0, ByVal VarPtr(REPF2)
End If
End Property

Public Property Get SelCharOffset() As Variant
Attribute SelCharOffset.VB_Description = "Returns/sets a value that determines whether text appears on the baseline (normal), as a superscript above the baseline, or as a subscript below the baseline."
Attribute SelCharOffset.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_OFFSET
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_OFFSET) <> 0 Then
        SelCharOffset = UserControl.ScaleY(.YOffset, vbTwips, vbContainerSize)
    Else
        SelCharOffset = Null
    End If
    End With
End If
End Property

Public Property Let SelCharOffset(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_OFFSET
    Select Case VarType(Value)
        Case vbSingle, vbDouble, vbLong, vbInteger, vbByte
            .YOffset = UserControl.ScaleY(Value, vbContainerSize, vbTwips)
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelColor() As Variant
Attribute SelColor.VB_Description = "Returns/sets a value that determines the color of the currently selected text."
Attribute SelColor.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_COLOR
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_COLOR) <> 0 Then
        SelColor = .TextColor
    Else
        SelColor = Null
    End If
    End With
End If
End Property

Public Property Let SelColor(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_COLOR
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            .TextColor = WinColor(Value)
        Case vbDouble, vbSingle
            .TextColor = WinColor(CLng(Value))
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelBkColor() As Variant
Attribute SelBkColor.VB_Description = "Returns/sets a value that determines the background color of the currently selected text."
Attribute SelBkColor.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_BACKCOLOR
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_BACKCOLOR) <> 0 Then
        SelBkColor = .BackColor
    Else
        SelBkColor = Null
    End If
    End With
End If
End Property

Public Property Let SelBkColor(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_BACKCOLOR
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            .BackColor = WinColor(Value)
        Case vbDouble, vbSingle
            .BackColor = WinColor(CLng(Value))
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelFontName() As Variant
Attribute SelFontName.VB_Description = "Returns/sets the font used to display the currently selected text."
Attribute SelFontName.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_FACE
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_FACE) <> 0 Then
        SelFontName = Left$(.FaceName(), InStr(CStr(.FaceName()) & vbNullChar, vbNullChar) - 1)
    Else
        SelFontName = Null
    End If
    End With
End If
End Property

Public Property Let SelFontName(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_FACE
    Select Case VarType(Value)
        Case vbString
            Dim Length As Long, FontB() As Byte
            If Len(Value) > LF_FACESIZE Then
                Length = LF_FACESIZE * 2
            Else
                Length = LenB(Value)
            End If
            If Length > 0 Then
                FontB() = Value
                CopyMemory .FaceName(0), FontB(0), Length
            End If
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelFontSize() As Variant
Attribute SelFontSize.VB_Description = "Returns/sets a value that specifies the size of the font used to display the currently selected text."
Attribute SelFontSize.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_SIZE
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_SIZE) <> 0 Then
        SelFontSize = CSng((.YHeight * 72) / 1440)
    Else
        SelFontSize = Null
    End If
    End With
End If
End Property

Public Property Let SelFontSize(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_SIZE
    Select Case VarType(Value)
        Case vbCurrency, vbSingle, vbDouble, vbLong, vbInteger, vbByte
            .YHeight = (Value * 1440) / 72
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelFontCharset() As Variant
Attribute SelFontCharset.VB_Description = "Returns/sets a value that specifies the charset of the font used to display the currently selected text."
Attribute SelFontCharset.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_CHARSET
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_CHARSET) <> 0 Then
        SelFontCharset = CInt(.Charset)
    Else
        SelFontCharset = Null
    End If
    End With
End If
End Property

Public Property Let SelFontCharset(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_CHARSET
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            .Charset = CByte(Value And &HFF)
        Case vbDouble, vbSingle
            .Charset = CByte(CLng(Value) And &HFF)
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelProtected() As Variant
Attribute SelProtected.VB_Description = "Returns/sets a value that determines if the selected text is protected against editing."
Attribute SelProtected.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_PROTECTED
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_PROTECTED) <> 0 Then
        SelProtected = CBool((.dwEffects And CFE_PROTECTED) = CFE_PROTECTED)
    Else
        SelProtected = Null
    End If
    End With
End If
End Property

Public Property Let SelProtected(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_PROTECTED
    Select Case VarType(Value)
        Case vbBoolean
            If Value = True Then .dwEffects = CFE_PROTECTED Else .dwEffects = 0
        Case vbNull
            .dwEffects = CFE_PROTECTED
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            If CBool(Value) = True Then .dwEffects = CFE_PROTECTED Else .dwEffects = 0
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelIndent() As Variant
Attribute SelIndent.VB_Description = "Returns/sets the distance between the left edge of the rich text box control and the left edge of the text that is selected or added at the current insertion point."
Attribute SelIndent.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_STARTINDENT
    If (SendMessage(RichTextBoxHandle, EM_GETPARAFORMAT, 0, ByVal VarPtr(REPF2)) And PFM_STARTINDENT) <> 0 Then
        SelIndent = UserControl.ScaleX(.DXStartIndent, vbTwips, vbContainerSize)
    Else
        SelIndent = Null
    End If
    End With
End If
End Property

Public Property Let SelIndent(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_STARTINDENT
    Select Case VarType(Value)
        Case vbSingle, vbDouble, vbLong, vbInteger, vbByte
            .DXStartIndent = UserControl.ScaleX(Value, vbContainerSize, vbTwips)
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETPARAFORMAT, 0, ByVal VarPtr(REPF2)
End If
End Property

Public Property Get SelRightIndent() As Variant
Attribute SelRightIndent.VB_Description = "Returns/sets the distance between the right edge of the rich text box control and the right edge of the text that is selected or added at the current insertion point."
Attribute SelRightIndent.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_RIGHTINDENT
    If (SendMessage(RichTextBoxHandle, EM_GETPARAFORMAT, 0, ByVal VarPtr(REPF2)) And PFM_RIGHTINDENT) <> 0 Then
        SelRightIndent = UserControl.ScaleX(.DXRightIndent, vbTwips, vbContainerSize)
    Else
        SelRightIndent = Null
    End If
    End With
End If
End Property

Public Property Let SelRightIndent(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_RIGHTINDENT
    Select Case VarType(Value)
        Case vbSingle, vbDouble, vbLong, vbInteger, vbByte
            .DXRightIndent = UserControl.ScaleX(Value, vbContainerSize, vbTwips)
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETPARAFORMAT, 0, ByVal VarPtr(REPF2)
End If
End Property

Public Property Get SelHangingIndent() As Variant
Attribute SelHangingIndent.VB_Description = "Returns/sets the distance between the left edge of the first line of text in the selected paragraph(s) and the left edge of subsequent lines of text in the same paragraph(s)."
Attribute SelHangingIndent.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_OFFSET
    If (SendMessage(RichTextBoxHandle, EM_GETPARAFORMAT, 0, ByVal VarPtr(REPF2)) And PFM_OFFSET) <> 0 Then
        SelHangingIndent = UserControl.ScaleX(.DXOffset, vbTwips, vbContainerSize)
    Else
        SelHangingIndent = Null
    End If
    End With
End If
End Property

Public Property Let SelHangingIndent(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_OFFSET
    Select Case VarType(Value)
        Case vbSingle, vbDouble, vbLong, vbInteger, vbByte
            .DXOffset = UserControl.ScaleX(Value, vbContainerSize, vbTwips)
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETPARAFORMAT, 0, ByVal VarPtr(REPF2)
End If
End Property

Public Property Get SelVisible() As Variant
Attribute SelVisible.VB_Description = "Returns/sets a value that determines if the selected text is visible or hidden."
Attribute SelVisible.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_HIDDEN
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_HIDDEN) <> 0 Then
        SelVisible = Not CBool((.dwEffects And CFE_HIDDEN) = CFE_HIDDEN)
    Else
        SelVisible = Null
    End If
    End With
End If
End Property

Public Property Let SelVisible(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_HIDDEN
    Select Case VarType(Value)
        Case vbBoolean
            If Value = False Then .dwEffects = CFE_HIDDEN Else .dwEffects = 0
        Case vbNull
            .dwEffects = 0
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            If CBool(Value) = False Then .dwEffects = CFE_HIDDEN Else .dwEffects = 0
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelLink() As Variant
Attribute SelLink.VB_Description = "Returns/sets a value that determines if the selected text is marked as hyperlink or not."
Attribute SelLink.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_LINK
    If (SendMessage(RichTextBoxHandle, EM_GETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)) And CFM_LINK) <> 0 Then
        SelLink = CBool((.dwEffects And CFE_LINK) = CFE_LINK)
    Else
        SelLink = Null
    End If
    End With
End If
End Property

Public Property Let SelLink(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim RECF2 As RECHARFORMAT2
    With RECF2
    .cbSize = LenB(RECF2)
    .dwMask = CFM_LINK
    Select Case VarType(Value)
        Case vbBoolean
            If Value = True Then .dwEffects = CFE_LINK Else .dwEffects = 0
        Case vbNull
            .dwEffects = 0
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            If CBool(Value) = True Then .dwEffects = CFE_LINK Else .dwEffects = 0
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETCHARFORMAT, SCF_SELECTION, ByVal VarPtr(RECF2)
End If
End Property

Public Property Get SelTabCount() As Variant
Attribute SelTabCount.VB_Description = "Returns/sets the number of tabs in the current selected text."
Attribute SelTabCount.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_TABSTOPS
    If (SendMessage(RichTextBoxHandle, EM_GETPARAFORMAT, 0, ByVal VarPtr(REPF2)) And PFM_TABSTOPS) <> 0 Then
        SelTabCount = .TabCount
    Else
        SelTabCount = Null
    End If
    End With
End If
End Property

Public Property Let SelTabCount(ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_TABSTOPS
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            If Value >= 0 And Value <= MAX_TAB_STOPS Then
                .TabCount = Value
            Else
                Err.Raise 380
            End If
        Case vbDouble, vbSingle
            If CLng(Value) >= 0 And CLng(Value) <= MAX_TAB_STOPS Then
                .TabCount = CLng(Value)
            Else
                Err.Raise 380
            End If
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETPARAFORMAT, 0, ByVal VarPtr(REPF2)
End If
End Property

Public Property Get SelTabs(ByVal Element As Integer) As Variant
Attribute SelTabs.VB_Description = "Returns/sets the absolute tab positions of the currently selected text."
Attribute SelTabs.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_TABSTOPS
    If (SendMessage(RichTextBoxHandle, EM_GETPARAFORMAT, 0, ByVal VarPtr(REPF2)) And PFM_TABSTOPS) <> 0 Then
        If Element >= 0 And .TabCount > 0 And Element < .TabCount And Element <= (MAX_TAB_STOPS - 1) Then
            SelTabs = UserControl.ScaleX(.Tabs(Element), vbTwips, vbContainerSize)
        Else
            Err.Raise 381
        End If
    Else
        SelTabs = Null
    End If
    End With
End If
End Property

Public Property Let SelTabs(ByVal Element As Integer, ByVal Value As Variant)
If RichTextBoxHandle <> 0 Then
    Dim REPF2 As REPARAFORMAT2
    With REPF2
    .cbSize = LenB(REPF2)
    .dwMask = PFM_TABSTOPS
    Select Case VarType(Value)
        Case vbSingle, vbDouble, vbLong, vbInteger, vbByte
            SendMessage RichTextBoxHandle, EM_GETPARAFORMAT, 0, ByVal VarPtr(REPF2)
            If Element >= 0 And .TabCount > 0 And Element < .TabCount And Element <= (MAX_TAB_STOPS - 1) Then
                .dwMask = PFM_TABSTOPS
                .Tabs(Element) = CLng(UserControl.ScaleX(Value, vbContainerSize, vbTwips))
            Else
                Err.Raise 381
            End If
        Case Else
            Err.Raise 13
    End Select
    End With
    SendMessage RichTextBoxHandle, EM_SETPARAFORMAT, 0, ByVal VarPtr(REPF2)
End If
End Property

Public Sub SelPrint(ByVal hDC As Long, Optional ByVal CallStartEndDoc As Boolean = True, Optional ByVal DocName As String = "RICHTEXT", Optional ByVal LeftMargin As Long, Optional ByVal TopMargin As Long, Optional ByVal RightMargin As Long, Optional ByVal BottomMargin As Long)
Attribute SelPrint.VB_Description = "Sends formatted text in a rich text box control to a device for printing."
If RichTextBoxHandle <> 0 And hDC <> 0 Then
    Dim RECR As RECHARRANGE, Length As Long
    If SendMessage(RichTextBoxHandle, EM_SELECTIONTYPE, 0, ByVal 0&) = RtfSelTypeEmpty Then
        RECR.Min = 0
        RECR.Max = -1
        Dim REGTLEX As REGETTEXTLENGTHEX
        REGTLEX.Flags = GTL_PRECISE Or GTL_NUMCHARS
        REGTLEX.CodePage = CP_UNICODE
        Length = SendMessage(RichTextBoxHandle, EM_GETTEXTLENGTHEX, VarPtr(REGTLEX), ByVal 0&)
    Else
        SendMessage RichTextBoxHandle, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
        Length = RECR.Max - RECR.Min
    End If
    Call CreatePrintJob(RECR.Min, RECR.Max, Length, hDC, CallStartEndDoc, DocName, LeftMargin, TopMargin, RightMargin, BottomMargin)
End If
End Sub

Public Sub PrintDoc(ByVal hDC As Long, Optional ByVal CallStartEndDoc As Boolean = True, Optional ByVal DocName As String = "RICHTEXT", Optional ByVal LeftMargin As Long, Optional ByVal TopMargin As Long, Optional ByVal RightMargin As Long, Optional ByVal BottomMargin As Long)
Attribute PrintDoc.VB_Description = "Sends formatted text in a rich text box control to a device for printing."
If RichTextBoxHandle <> 0 And hDC <> 0 Then
    Dim Length As Long, REGTLEX As REGETTEXTLENGTHEX
    REGTLEX.Flags = GTL_PRECISE Or GTL_NUMCHARS
    REGTLEX.CodePage = CP_UNICODE
    Length = SendMessage(RichTextBoxHandle, EM_GETTEXTLENGTHEX, VarPtr(REGTLEX), ByVal 0&)
    Call CreatePrintJob(0, -1, Length, hDC, CallStartEndDoc, DocName, LeftMargin, TopMargin, RightMargin, BottomMargin)
End If
End Sub

Public Sub SaveFile(ByVal FileName As String, Optional ByVal Format As RtfLoadSaveFormatConstants = RtfLoadSaveFormatRTF, Optional ByVal SelectionOnly As Boolean)
Attribute SaveFile.VB_Description = "Saves the contents of a rich text box control to a file."
If FileName = vbNullString Then Exit Sub
Dim Flags As Long
If SelectionOnly = True Then Flags = SFF_SELECTION
Select Case Format
    Case RtfLoadSaveFormatRTF
        Flags = Flags Or SF_RTF
    Case RtfLoadSaveFormatText
        Flags = Flags Or SF_TEXT
    Case RtfLoadSaveFormatUnicodeText
        Flags = Flags Or SF_TEXT Or SF_UNICODE
End Select
StreamFileOut FileName, Flags
End Sub

Public Sub LoadFile(ByVal FileName As String, Optional ByVal Format As RtfLoadSaveFormatConstants = RtfLoadSaveFormatRTF, Optional ByVal SelectionOnly As Boolean)
Attribute LoadFile.VB_Description = "Loads an .RTF file or text file into a rich text box control."
If FileName = vbNullString Then Exit Sub
If FileExists(FileName) = True Then
    Dim Flags As Long
    If SelectionOnly = True Then Flags = SFF_SELECTION
    Select Case Format
        Case RtfLoadSaveFormatRTF
            Flags = Flags Or SF_RTF
        Case RtfLoadSaveFormatText
            Flags = Flags Or SF_TEXT
        Case RtfLoadSaveFormatUnicodeText
            Flags = Flags Or SF_TEXT Or SF_UNICODE
    End Select
    StreamFileIn FileName, Flags
Else
    If RichTextBoxDesignMode = True Then
        MsgBox "The specified file name cannot be accessed or is invalid.", vbCritical + vbOKOnly
        Exit Sub
    Else
        Err.Raise Number:=75, Description:="The specified file name cannot be accessed or is invalid"
    End If
End If
End Sub

Public Function GetLine(ByVal LineNumber As Long) As String
Attribute GetLine.VB_Description = "Retrieves the text of the specified line. A value of 0 indicates that the text of the current line number (the line that contains the caret) will be retrieved."
If LineNumber < 0 Then Err.Raise 380
If RichTextBoxHandle <> 0 Then
    Dim FirstCharPos As Long, Length As Long
    FirstCharPos = SendMessage(RichTextBoxHandle, EM_LINEINDEX, LineNumber - 1, ByVal 0&)
    If FirstCharPos > -1 Then
        Length = SendMessage(RichTextBoxHandle, EM_LINELENGTH, FirstCharPos, ByVal 0&)
        If Length > 0 Then
            Dim Buffer As String
            Buffer = ChrW(Length) & String$(Length - 1, vbNullChar)
            If LineNumber > 0 Then
                If SendMessage(RichTextBoxHandle, EM_GETLINE, LineNumber - 1, ByVal StrPtr(Buffer)) > 0 Then GetLine = Buffer
            Else
                If SendMessage(RichTextBoxHandle, EM_GETLINE, SendMessage(RichTextBoxHandle, EM_EXLINEFROMCHAR, 0, ByVal FirstCharPos), ByVal StrPtr(Buffer)) > 0 Then GetLine = Buffer
            End If
        End If
    Else
        Err.Raise 380
    End If
End If
End Function

Public Function GetLineCount() As Long
Attribute GetLineCount.VB_Description = "Gets the number of lines."
If RichTextBoxHandle <> 0 Then GetLineCount = SendMessage(RichTextBoxHandle, EM_GETLINECOUNT, 0, ByVal 0&)
End Function

Public Sub ScrollToLine(ByVal LineNumber As Long)
Attribute ScrollToLine.VB_Description = "Scrolls to ensure the specified line is visible."
If LineNumber < 0 Then Err.Raise 380
If RichTextBoxHandle <> 0 Then
    If SendMessage(RichTextBoxHandle, EM_LINEINDEX, LineNumber - 1, ByVal 0&) > -1 Then
        Dim LineIndex As Long
        LineIndex = SendMessage(RichTextBoxHandle, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&)
        SendMessage RichTextBoxHandle, EM_LINESCROLL, 0, ByVal CLng((LineNumber - 1) - LineIndex)
    Else
        Err.Raise 380
    End If
End If
End Sub

Public Sub ScrollToCaret()
Attribute ScrollToCaret.VB_Description = "Scrolls the caret into view."
If RichTextBoxHandle <> 0 Then
    ' RichEdit bug that EM_SCROLLCARET works only when the control has the focus.
    ' There is a workaround though to temporarily show the selection and hide again.
    If PropHideSelection = True Then
        SendMessage RichTextBoxHandle, EM_HIDESELECTION, 0, ByVal 0&
        SendMessage RichTextBoxHandle, EM_SCROLLCARET, 0, ByVal 0&
        SendMessage RichTextBoxHandle, EM_HIDESELECTION, 1, ByVal 0&
    Else
        SendMessage RichTextBoxHandle, EM_SCROLLCARET, 0, ByVal 0&
    End If
End If
End Sub

Public Function CharFromPos(ByVal X As Single, ByVal Y As Single) As Long
Attribute CharFromPos.VB_Description = "Returns the character index closest to a specified point."
Dim P As POINTAPI
P.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
P.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
If RichTextBoxHandle <> 0 Then CharFromPos = SendMessage(RichTextBoxHandle, EM_CHARFROMPOS, 0, ByVal VarPtr(P))
End Function

Public Function GetLineFromChar(ByVal CharIndex As Long) As Long
Attribute GetLineFromChar.VB_Description = "Gets the line number that contains the specified character index. A character index of -1 retrieves either the current line or the beginning of the current selection."
If CharIndex < -1 Then Err.Raise 380
If RichTextBoxHandle <> 0 Then GetLineFromChar = SendMessage(RichTextBoxHandle, EM_EXLINEFROMCHAR, 0, ByVal CharIndex) + 1
End Function

Public Function GetSelType() As Integer
Attribute GetSelType.VB_Description = "Determines the selection type."
If RichTextBoxHandle <> 0 Then GetSelType = SendMessage(RichTextBoxHandle, EM_SELECTIONTYPE, 0, ByVal 0&)
End Function

Public Property Get LeftMargin() As Single
Attribute LeftMargin.VB_Description = "Returns/sets the widths of the left margin."
Attribute LeftMargin.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then LeftMargin = UserControl.ScaleX(LoWord(SendMessage(RichTextBoxHandle, EM_GETMARGINS, 0, ByVal 0&)), vbPixels, vbContainerSize)
End Property

Public Property Let LeftMargin(ByVal Value As Single)
If Value = EC_USEFONTINFO Or Value = -1 Then
    If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_SETMARGINS, EC_LEFTMARGIN Or EC_USEFONTINFO, ByVal 0&
Else
    If Value < 0 Then Err.Raise 380
    Dim IntValue As Integer
    IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_SETMARGINS, EC_LEFTMARGIN, ByVal MakeDWord(IntValue, 0)
End If
UserControl.PropertyChanged "LeftMargin"
End Property

Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Returns/sets the widths of the right margin."
Attribute RightMargin.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then RightMargin = UserControl.ScaleX(HiWord(SendMessage(RichTextBoxHandle, EM_GETMARGINS, 0, ByVal 0&)), vbPixels, vbContainerSize)
End Property

Public Property Let RightMargin(ByVal Value As Single)
If Value = EC_USEFONTINFO Or Value = -1 Then
    If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_SETMARGINS, EC_RIGHTMARGIN Or EC_USEFONTINFO, ByVal 0&
Else
    If Value < 0 Then Err.Raise 380
    Dim IntValue As Integer
    IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_SETMARGINS, EC_RIGHTMARGIN, ByVal MakeDWord(0, IntValue)
End If
UserControl.PropertyChanged "RightMargin"
End Property

Public Property Get ZoomFactor() As Double
Attribute ZoomFactor.VB_Description = "Returns/sets the current zoom factor."
Attribute ZoomFactor.VB_MemberFlags = "400"
If RichTextBoxHandle <> 0 Then
    Dim Numerator As Long, Denominator As Long
    SendMessage RichTextBoxHandle, EM_GETZOOM, VarPtr(Numerator), ByVal VarPtr(Denominator)
    If Numerator = 0 And Denominator = 0 Then
        ZoomFactor = 1
    Else
        ZoomFactor = Numerator / Denominator
    End If
End If
End Property

Public Property Let ZoomFactor(ByVal Value As Double)
Dim Numerator As Long, Denominator As Long
If Value = 1 Then
    Numerator = 0
    Denominator = 0
Else
    Numerator = 1000 * Value
    Denominator = 1000
End If
If RichTextBoxHandle <> 0 Then
    If SendMessage(RichTextBoxHandle, EM_SETZOOM, Numerator, ByVal Denominator) = 0 Then Err.Raise 380
End If
UserControl.PropertyChanged "ZoomFactor"
End Property

Public Function GetOLEInterface() As IUnknown
Attribute GetOLEInterface.VB_Description = "Retrieves an IRichEditOle object that a client can use to access the COM functionality."
If RichTextBoxHandle <> 0 Then SendMessage RichTextBoxHandle, EM_GETOLEINTERFACE, 0, GetOLEInterface
End Function

Public Sub OLEObjectsAdd(ByVal LpOleObject As Long)
Attribute OLEObjectsAdd.VB_Description = "Inserts an OLE object into a rich text box control."
If RichTextBoxHandle <> 0 Then
    Dim OLEInstance As OLEGuids.IRichEditOle
    Set OLEInstance = Me.GetOLEInterface
    If Not OLEInstance Is Nothing Then
        Dim PropOleObject As OLEGuids.IOleObject, PropClientSite As OLEGuids.IOleClientSite, PropStorage As OLEGuids.IStorage
        Set PropOleObject = PtrToObj(LpOleObject)
        Set PropClientSite = OLEInstance.GetClientSite
        StgCreateDocFile 0, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE Or STGM_DELETEONRELEASE, 0, PropStorage
        Const IID_IOleObject As String = "{00000112-0000-0000-C000-000000000046}"
        Dim IID As OLEGuids.OLECLSID
        CLSIDFromString StrPtr(IID_IOleObject), IID
        If Not PropOleObject Is Nothing Then
            OleSetContainedObject PropOleObject, 1
            Dim REOBJ As REOBJECT
            With REOBJ
            .cbStruct = LenB(REOBJ)
            LSet .riid = IID
            .dvAspect = DVASPECT_CONTENT
            .CharPos = REO_CP_SELECTION
            .dwFlags = REO_DYNAMICSIZE Or REO_RESIZABLE Or REO_BELOWBASELINE
            .Size.CX = 0
            .Size.CY = 0
            .dwUser = 0
            Set .pStorage = PropStorage
            Set .pOleSite = PropClientSite
            Set .pOleObject = PropOleObject
            End With
            OLEInstance.InsertObject REOBJ
        End If
    End If
End If
End Sub

Public Sub OLEObjectsAddFromFile(ByVal FileName As String, Optional ByVal LinkToFile As Boolean)
Attribute OLEObjectsAddFromFile.VB_Description = "Inserts an OLE object (from file) into a rich text box control."
If RichTextBoxHandle <> 0 Then
    Dim OLEInstance As OLEGuids.IRichEditOle
    Set OLEInstance = Me.GetOLEInterface
    If Not OLEInstance Is Nothing Then
        Dim PropOleObject As OLEGuids.IOleObject, PropClientSite As OLEGuids.IOleClientSite, PropStorage As OLEGuids.IStorage
        Set PropClientSite = OLEInstance.GetClientSite
        StgCreateDocFile 0, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE Or STGM_DELETEONRELEASE, 0, PropStorage
        Const IID_IOleObject As String = "{00000112-0000-0000-C000-000000000046}"
        Dim IID As OLEGuids.OLECLSID
        CLSIDFromString StrPtr(IID_IOleObject), IID
        Dim IID_NULL As OLEGuids.OLECLSID
        Const STG_E_FILENOTFOUND As Long = &H80030002
        If LinkToFile = False Then
            If OleCreateFromFile(IID_NULL, StrPtr(FileName), IID, OLERENDER_DRAW, 0, PropClientSite, PropStorage, PropOleObject) = STG_E_FILENOTFOUND Then Err.Raise 53
        Else
            If OleCreateLinkToFile(StrPtr(FileName), IID, OLERENDER_DRAW, 0, PropClientSite, PropStorage, PropOleObject) = STG_E_FILENOTFOUND Then Err.Raise 53
        End If
        If Not PropOleObject Is Nothing Then
            OleSetContainedObject PropOleObject, 1
            Dim REOBJ As REOBJECT
            With REOBJ
            .cbStruct = LenB(REOBJ)
            LSet .riid = IID
            .dvAspect = DVASPECT_CONTENT
            .CharPos = REO_CP_SELECTION
            .dwFlags = REO_DYNAMICSIZE Or REO_RESIZABLE Or REO_BELOWBASELINE
            .Size.CX = 0
            .Size.CY = 0
            .dwUser = 0
            Set .pStorage = PropStorage
            Set .pOleSite = PropClientSite
            Set .pOleObject = PropOleObject
            End With
            OLEInstance.InsertObject REOBJ
        End If
    End If
End If
End Sub

Public Sub OLEObjectsAddFromPicture(ByVal Picture As IPictureDisp, Optional ByVal Format As Variant)
Attribute OLEObjectsAddFromPicture.VB_Description = "Inserts an OLE object (from picture object) into a rich text box control."
If Not Picture Is Nothing Then
    If Picture.Handle <> 0 Then
        Select Case Picture.Type
            Case vbPicTypeBitmap, vbPicTypeMetafile, vbPicTypeEMetafile
                If IsMissing(Format) = True Then
                    RichTextBoxDataObjectFormat = Empty
                Else
                    Select Case VarType(Format)
                        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
                            Select Case Format
                                Case vbCFBitmap
                                    RichTextBoxDataObjectFormat = vbCFBitmap
                                Case vbCFDIB
                                    RichTextBoxDataObjectFormat = vbCFDIB
                                Case vbCFMetafile
                                    RichTextBoxDataObjectFormat = vbCFMetafile
                                Case vbCFEMetafile
                                    RichTextBoxDataObjectFormat = vbCFEMetafile
                                Case Else
                                    Err.Raise Number:=461, Description:="Specified format doesn't match format of data"
                            End Select
                        Case Else
                            Err.Raise 13
                    End Select
                End If
                Set RichTextBoxDataObjectValue = Picture
                ' Start of a fake OLE drag/drop event.
                ' Actual picture insertion happens in the OLEStartDrag event.
                UserControl.OLEDrag
                RichTextBoxDataObjectValue = Empty
                RichTextBoxDataObjectFormat = Empty
            Case Else
                Err.Raise 380
        End Select
    End If
End If
End Sub

Public Function OLEObjectsGet(ByVal IndexObj As Long, Optional ByVal CharPos As Long) As Long
Attribute OLEObjectsGet.VB_Description = "Retrieves an OLE object in a rich text box control."
If RichTextBoxHandle <> 0 Then
    Dim OLEInstance As OLEGuids.IRichEditOle
    Set OLEInstance = Me.GetOLEInterface
    If Not OLEInstance Is Nothing Then
        Dim REOBJ As REOBJECT
        REOBJ.cbStruct = LenB(REOBJ)
        If IndexObj = REO_IOB_USE_CP Then REOBJ.CharPos = CharPos
        OLEInstance.GetObject IndexObj, REOBJ, REO_GETOBJ_POLEOBJ
        OLEObjectsGet = ObjPtr(REOBJ.pOleObject)
    End If
End If
End Function

Public Function OLEObjectsCount() As Long
Attribute OLEObjectsCount.VB_Description = "Returns the number of OLE objects currently contained in a rich text box control."
If RichTextBoxHandle <> 0 Then
    Dim OLEInstance As OLEGuids.IRichEditOle
    Set OLEInstance = Me.GetOLEInterface
    If Not OLEInstance Is Nothing Then OLEObjectsCount = OLEInstance.GetObjectCount
End If
End Function

Private Function StreamStringOut(ByRef Value As String, ByVal Flags As Long) As Long
If RichTextBoxHandle <> 0 Then
    Dim REEDSTR As REEDITSTREAM
    With REEDSTR
    .dwCookie = 0
    .dwError = 0
    .lpfnCallback = ProcPtr(AddressOf RtfStreamCallbackStringOut)
    End With
    StreamStringOut = SendMessage(RichTextBoxHandle, EM_STREAMOUT, Flags, ByVal VarPtr(REEDSTR))
    If (Flags And SF_UNICODE) = 0 Then
        Value = StrConv(RtfStreamStringOut(), vbUnicode)
    Else
        Value = RtfStreamStringOut()
    End If
End If
End Function

Private Function StreamStringIn(ByRef Value As String, ByVal Flags As Long) As Long
If RichTextBoxHandle <> 0 Then
    Dim REEDSTR As REEDITSTREAM
    With REEDSTR
    .dwCookie = 0
    .dwError = 0
    .lpfnCallback = ProcPtr(AddressOf RtfStreamCallbackStringIn)
    End With
    If (Flags And SF_UNICODE) = 0 Then
        If Len(Value) <> LenB(Value) Then
            Call RtfStreamStringIn(StrConv(Value, vbFromUnicode))
        Else
            Call RtfStreamStringIn(Value)
        End If
    Else
        Call RtfStreamStringIn(Value)
    End If
    StreamStringIn = SendMessage(RichTextBoxHandle, EM_STREAMIN, Flags, ByVal VarPtr(REEDSTR))
    Call RtfStreamStringInCleanUp
End If
End Function

Private Function StreamFileOut(ByVal FileName As String, ByVal Flags As Long) As Long
If RichTextBoxHandle <> 0 Then
    If Left$(FileName, 2) = "\\" Then FileName = "UNC\" & Mid$(FileName, 3)
    Dim hFile As Long
    hFile = CreateFile(StrPtr("\\?\" & FileName), GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        Dim REEDSTR As REEDITSTREAM
        With REEDSTR
        .dwCookie = hFile
        .dwError = 0
        .lpfnCallback = ProcPtr(AddressOf RtfStreamCallbackFileOut)
        End With
        If (Flags And SF_UNICODE) <> 0 Then
            Dim B(0 To 1) As Byte ' UTF-16 BOM
            B(0) = &HFF
            B(1) = &HFE
            WriteFile hFile, VarPtr(B(0)), 2, 0, 0
        End If
        StreamFileOut = SendMessage(RichTextBoxHandle, EM_STREAMOUT, Flags, ByVal VarPtr(REEDSTR))
        CloseHandle hFile
    End If
End If
End Function

Private Function StreamFileIn(ByVal FileName As String, ByVal Flags As Long) As Long
If RichTextBoxHandle <> 0 Then
    If Left$(FileName, 2) = "\\" Then FileName = "UNC\" & Mid$(FileName, 3)
    Dim hFile As Long
    hFile = CreateFile(StrPtr("\\?\" & FileName), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        Dim REEDSTR As REEDITSTREAM
        With REEDSTR
        .dwCookie = hFile
        .dwError = 0
        .lpfnCallback = ProcPtr(AddressOf RtfStreamCallbackFileIn)
        End With
        If (Flags And SF_UNICODE) <> 0 Then
            Dim B(0 To 1) As Byte, dwRead As Long
            ReadFile hFile, VarPtr(B(0)), 2, dwRead, 0
            If B(0) = &HFF And B(1) = &HFE Then ' UTF-16 BOM
            ElseIf dwRead > 0 Then
                SetFilePointer hFile, 0, ByVal 0, FILE_BEGIN
            End If
        End If
        StreamFileIn = SendMessage(RichTextBoxHandle, EM_STREAMIN, Flags, ByVal VarPtr(REEDSTR))
        CloseHandle hFile
    End If
End If
End Function

Private Sub CreatePrintJob(ByVal Min As Long, ByVal Max As Long, ByVal Length As Long, ByVal hDC As Long, ByVal CallStartEndDoc As Boolean, ByVal DocName As String, ByVal LeftMargin As Long, ByVal TopMargin As Long, ByVal RightMargin As Long, ByVal BottomMargin As Long)
If RichTextBoxHandle <> 0 And hDC <> 0 Then
    Dim REFR As REFORMATRANGE
    With REFR
    .hDC = hDC
    .hDCTarget = hDC
    .CharRange.Min = Min
    .CharRange.Max = Max
    Dim IsPrinterDC As Boolean, PhysCX As Long, PhysCY As Long, PhysOffsetCX As Long, PhysOffsetCY As Long
    IsPrinterDC = CBool(GetDeviceCaps(hDC, PHYSICALWIDTH) > 0 And GetDeviceCaps(hDC, PHYSICALHEIGHT) > 0)
    If IsPrinterDC = True Then
        PhysCX = MulDiv(GetDeviceCaps(hDC, PHYSICALWIDTH), 1440, GetDeviceCaps(hDC, LOGPIXELSX))
        PhysCY = MulDiv(GetDeviceCaps(hDC, PHYSICALHEIGHT), 1440, GetDeviceCaps(hDC, LOGPIXELSY))
        PhysOffsetCX = MulDiv(GetDeviceCaps(hDC, PHYSICALOFFSETX), 1440, GetDeviceCaps(hDC, LOGPIXELSX))
        PhysOffsetCY = MulDiv(GetDeviceCaps(hDC, PHYSICALOFFSETY), 1440, GetDeviceCaps(hDC, LOGPIXELSY))
    Else
        Dim hDCScreen As Long
        hDCScreen = GetDC(0)
        If hDCScreen <> 0 Then
            PhysCX = MulDiv(GetDeviceCaps(hDCScreen, HORZRES), 1440, GetDeviceCaps(hDCScreen, LOGPIXELSX))
            PhysCY = MulDiv(GetDeviceCaps(hDCScreen, VERTRES), 1440, GetDeviceCaps(hDCScreen, LOGPIXELSY))
            ReleaseDC 0, hDCScreen
        End If
    End If
    With .RCPage
    .Left = 0
    .Top = 0
    .Right = PhysCX
    .Bottom = PhysCY
    End With
    With .RC
    .Left = LeftMargin - PhysOffsetCX
    .Top = TopMargin - PhysOffsetCY
    .Right = (PhysCX - RightMargin) + PhysOffsetCX
    .Bottom = (PhysCY - BottomMargin) + PhysOffsetCY
    End With
    If IsPrinterDC = True Then
        If CallStartEndDoc = True Then
            Dim DI As DOCINFO
            With DI
            .cbSize = LenB(DI)
            .lpszDocName = StrPtr(DocName)
            .lpszOutput = 0
            .lpszDatatype = 0
            .fwType = 0
            End With
            StartDoc hDC, DI
        End If
        Dim NextCharPos As Long, Success As Boolean
        Do
            Success = CBool(StartPage(hDC) > 0)
            If Success = False Then Exit Do
            NextCharPos = SendMessage(RichTextBoxHandle, EM_FORMATRANGE, 1, ByVal VarPtr(REFR))
            Success = CBool(EndPage(hDC) > 0)
            If NextCharPos < Length Then .CharRange.Min = NextCharPos Else Exit Do
        Loop While Success = True
        SendMessage RichTextBoxHandle, EM_FORMATRANGE, 0, ByVal 0&
        If CallStartEndDoc = True Then
            If Success = True Then
                EndDoc hDC
            Else
                AbortDoc hDC
            End If
        End If
    Else
        SendMessage RichTextBoxHandle, EM_FORMATRANGE, 1, ByVal VarPtr(REFR)
        SendMessage RichTextBoxHandle, EM_FORMATRANGE, 0, ByVal 0&
    End If
    End With
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

Friend Sub FIRichEditOleCallback_GetNewStorage(ByRef RetVal As Long, ByRef ppStorage As OLEGuids.IStorage)
RetVal = StgCreateDocFile(0, STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE Or STGM_DELETEONRELEASE, 0, ppStorage)
End Sub

Friend Sub FIRichEditOleCallback_DeleteObject(ByVal LpOleObject As Long)
RaiseEvent OLEDeleteObject(LpOleObject)
End Sub

Friend Sub FIRichEditOleCallback_GetDragDropEffect(ByVal Drag As Boolean, ByVal KeyState As Long, ByRef dwEffect As Long)
If Drag = True Then
    RaiseEvent OLEStartDrag(dwEffect) ' AllowedEffects
Else
    Dim Pos As Long, P As POINTAPI
    Pos = GetMessagePos()
    P.X = Get_X_lParam(Pos)
    P.Y = Get_Y_lParam(Pos)
    ScreenToClient UserControl.hWnd, P
    RaiseEvent OLEGetDropEffect(dwEffect, GetMouseStateFromParam(KeyState), GetShiftStateFromParam(KeyState), UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
End If
End Sub

Friend Sub FIRichEditOleCallback_GetContextMenu(ByVal SelType As Integer, ByVal LpOleObject As Long, ByVal lpCharRange As Long, ByRef hMenu As Long)
Dim RECR As RECHARRANGE
CopyMemory RECR, ByVal lpCharRange, 8
RaiseEvent OLEGetContextMenu(SelType, LpOleObject, RECR.Min, RECR.Max, hMenu)
End Sub

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
            If KeyCode = vbKeyInsert Then
                If PropAllowOverType = False And PropOverTypeMode = False Then Exit Function
                If wMsg = WM_KEYDOWN Then PropOverTypeMode = Not PropOverTypeMode
            End If
            RichTextBoxCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If RichTextBoxCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(RichTextBoxCharCodeCache And &HFFFF&)
            RichTextBoxCharCodeCache = 0
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
        Call ComCtlsSetIMEMode(hWnd, RichTextBoxIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call ComCtlsSetIMEMode(hWnd, RichTextBoxIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_VSCROLL, WM_HSCROLL
        ' The notification codes EN_HSCROLL and EN_VSCROLL are not sent when clicking the scroll bar thumb itself.
        If LoWord(wParam) = SB_THUMBTRACK Then RaiseEvent Scroll
    Case WM_CONTEXTMENU
        If wParam = RichTextBoxHandle Then
            Dim P As POINTAPI, Handled As Boolean
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            If P.X = -1 And P.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(Handled, -1, -1)
            Else
                ScreenToClient RichTextBoxHandle, P
                RaiseEvent ContextMenu(Handled, UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
            End If
            If Handled = True Then Exit Function
        End If
    
    #If ImplementThemedBorder = True Then
    
    Case WM_THEMECHANGED, WM_STYLECHANGED, WM_ENABLE
        If wMsg = WM_THEMECHANGED Then RichTextBoxEnabledVisualStyles = EnabledVisualStyles()
        If PropBorderStyle = vbFixedSingle And PropVisualStyles = True Then
            If RichTextBoxEnabledVisualStyles = True Then SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
        End If
    Case WM_NCPAINT
        ' For some reason, Microsoft never updated its rich edit library after the release of Windows XP to make the rich edit control theme-aware.
        ' In order to support themes it is necessary to do a workaround.
        ' In addition the disabled and focused state will be handled.
        If PropBorderStyle = vbFixedSingle And PropVisualStyles = True And RichTextBoxEnabledVisualStyles = True Then
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
                    ' Printing the default non-client area ensures that the scrollbars are drawn, if any.
                    SendMessage hWnd, WM_PRINT, hDC, ByVal PRF_NONCLIENT
                    Dim BorderX As Long, BorderY As Long
                    Dim RC1 As RECT, RC2 As RECT
                    Const SM_CXEDGE As Long = 45
                    Const SM_CYEDGE As Long = 46
                    BorderX = GetSystemMetrics(SM_CXEDGE)
                    BorderY = GetSystemMetrics(SM_CYEDGE)
                    With UserControl
                    SetRect RC1, BorderX, BorderY, .ScaleWidth - BorderX, .ScaleHeight - BorderY
                    SetRect RC2, 0, 0, .ScaleWidth, .ScaleHeight
                    End With
                    ExcludeClipRect hDC, RC1.Left, RC1.Top, RC1.Right, RC1.Bottom
                    Dim dwStyle As Long
                    dwStyle = GetWindowLong(hWnd, GWL_STYLE)
                    Dim EditPart As Long, EditState As Long
                    If (dwStyle And WS_HSCROLL) = WS_HSCROLL Then
                        If (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                            EditPart = EP_EDITBORDER_HVSCROLL
                        Else
                            EditPart = EP_EDITBORDER_HSCROLL
                        End If
                    Else
                        If (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                            EditPart = EP_EDITBORDER_VSCROLL
                        Else
                            EditPart = EP_EDITBORDER_NOSCROLL
                        End If
                    End If
                    Dim Brush As Long
                    If Me.Enabled = False Then
                        EditState = EPSN_DISABLED
                        Brush = CreateSolidBrush(WinColor(vbButtonFace))
                    Else
                        If RichTextBoxFocused = True Then
                            EditState = EPSN_FOCUSED
                        ElseIf RichTextBoxMouseOver(0) = True Then
                            EditState = EPSN_HOT
                        Else
                            EditState = EPSN_NORMAL
                        End If
                        Brush = CreateSolidBrush(WinColor(PropBackColor))
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
    
    #End If
    
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    
    #If ImplementThemedBorder = True Then
    
    Case WM_SETFOCUS, WM_KILLFOCUS
        RichTextBoxFocused = CBool(wMsg = WM_SETFOCUS)
        If PropBorderStyle = vbFixedSingle And PropVisualStyles = True Then
            If RichTextBoxEnabledVisualStyles = True Then SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
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
                RichTextBoxIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                RichTextBoxIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                RichTextBoxIsClick = True
            Case WM_MOUSEMOVE
                If (RichTextBoxMouseOver(0) = False And PropBorderStyle = vbFixedSingle) Or (RichTextBoxMouseOver(1) = False And PropMouseTrack = True) Then
                    
                    #If ImplementThemedBorder = True Then
                    
                    If RichTextBoxMouseOver(0) = False And PropBorderStyle = vbFixedSingle Then
                        If RichTextBoxEnabledVisualStyles = True And PropVisualStyles = True Then
                            RichTextBoxMouseOver(0) = True
                            SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
                        End If
                    End If
                    
                    #End If
                    
                    If RichTextBoxMouseOver(1) = False And PropMouseTrack = True Then
                        RichTextBoxMouseOver(1) = True
                        RaiseEvent MouseEnter
                    End If
                    If RichTextBoxMouseOver(0) = True Or RichTextBoxMouseOver(1) = True Then Call ComCtlsRequestMouseLeave(hWnd)
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
                If RichTextBoxIsClick = True Then
                    RichTextBoxIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        
        #If ImplementThemedBorder = True Then
        
        If RichTextBoxMouseOver(0) = True Then
            RichTextBoxMouseOver(0) = False
            SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_DRAWFRAME
        End If
        
        #End If
        
        If RichTextBoxMouseOver(1) = True Then
            RichTextBoxMouseOver(1) = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        If HiWord(wParam) = 0 And lParam = 0 Then ' Alias for menu
            RaiseEvent OLEContextMenuClick(LoWord(wParam))
        ElseIf lParam <> 0 Then
            Select Case HiWord(wParam)
                Case EN_CHANGE
                    UserControl.PropertyChanged "Text"
                    UserControl.PropertyChanged "TextRTF"
                    On Error Resume Next
                    UserControl.Extender.DataChanged = True
                    On Error GoTo 0
                    RaiseEvent Change
                Case EN_MAXTEXT
                    RaiseEvent MaxText
                Case EN_HSCROLL, EN_VSCROLL
                    ' This notification code is also sent when a keyboard event causes a change in the view area.
                    RaiseEvent Scroll
            End Select
        End If
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = RichTextBoxHandle Then
            Select Case NM.Code
                Case EN_SELCHANGE
                    Dim NMENSC As NMENSELCHANGE
                    CopyMemory NMENSC, ByVal lParam, LenB(NMENSC)
                    With NMENSC
                    RaiseEvent SelChange(.SelType, .CharRange.Min, .CharRange.Max)
                    End With
                Case EN_DRAGDROPDONE
                    RaiseEvent OLECompleteDrag
                Case EN_LINK
                    Dim NMENL As NMENLINK
                    CopyMemory NMENL, ByVal lParam, LenB(NMENL)
                    With NMENL
                    RaiseEvent LinkEvent(.wMsg, .wParam, .lParam, .CharRange.Min, .CharRange.Max)
                    End With
                Case EN_PROTECTED
                    Dim NMENP As NMENPROTECTED
                    CopyMemory NMENP, ByVal lParam, LenB(NMENP)
                    Dim Allow As Boolean
                    With NMENP.CharRange
                    RaiseEvent ModifyProtected(Allow, .Min, .Max)
                    End With
                    If Allow = False Then
                        WindowProcUserControl = 1
                    Else
                        WindowProcUserControl = 0
                    End If
                    Exit Function
            End Select
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI RichTextBoxHandle
End Function
