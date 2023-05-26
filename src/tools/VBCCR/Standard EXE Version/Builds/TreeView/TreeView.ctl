VERSION 5.00
Begin VB.UserControl TreeView 
   Alignable       =   -1  'True
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "TreeView.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "TreeView.ctx":0049
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "TreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
#If False Then
Private TvwStyleTextOnly, TvwStylePictureText, TvwStylePlusMinusText, TvwStylePlusMinusPictureText, TvwStyleTreeLinesText, TvwStyleTreeLinesPictureText, TvwStyleTreeLinesPlusMinusText, TvwStyleTreeLinesPlusMinusPictureText
Private TvwLineStyleTreeLines, TvwLineStyleRootLines
Private TvwLabelEditAutomatic, TvwLabelEditManual, TvwLabelEditDisabled
Private TvwNodeRelationshipFirst, TvwNodeRelationshipLast, TvwNodeRelationshipNext, TvwNodeRelationshipPrevious, TvwNodeRelationshipChild
Private TvwSortOrderAscending, TvwSortOrderDescending
Private TvwSortTypeBinary, TvwSortTypeText
Private TvwMultiSelectNone, TvwMultiSelectAll, TvwMultiSelectVisibleOnly, TvwMultiSelectRestrictSiblings
Private TvwVisualThemeStandard, TvwVisualThemeExplorer
#End If
Public Enum TvwStyleConstants
TvwStyleTextOnly = 0
TvwStylePictureText = 1
TvwStylePlusMinusText = 2
TvwStylePlusMinusPictureText = 3
TvwStyleTreeLinesText = 4
TvwStyleTreeLinesPictureText = 5
TvwStyleTreeLinesPlusMinusText = 6
TvwStyleTreeLinesPlusMinusPictureText = 7
End Enum
Public Enum TvwLineStyleConstants
TvwLineStyleTreeLines = 0
TvwLineStyleRootLines = 1
End Enum
Public Enum TvwLabelEditConstants
TvwLabelEditAutomatic = 0
TvwLabelEditManual = 1
TvwLabelEditDisabled = 2
End Enum
Public Enum TvwNodeRelationshipConstants
TvwNodeRelationshipFirst = 0
TvwNodeRelationshipLast = 1
TvwNodeRelationshipNext = 2
TvwNodeRelationshipPrevious = 3
TvwNodeRelationshipChild = 4
End Enum
Public Enum TvwSortOrderConstants
TvwSortOrderAscending = 0
TvwSortOrderDescending = 1
End Enum
Public Enum TvwSortTypeConstants
TvwSortTypeBinary = 0
TvwSortTypeText = 1
End Enum
Public Enum TvwMultiSelectConstants
TvwMultiSelectNone = 0
TvwMultiSelectAll = 1
TvwMultiSelectVisibleOnly = 2
TvwMultiSelectRestrictSiblings = 3
End Enum
Public Enum TvwVisualThemeConstants
TvwVisualThemeStandard = 0
TvwVisualThemeExplorer = 1
End Enum
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type TVITEM
Mask As Long
hItem As Long
State As Long
StateMask As Long
pszText As Long
cchTextMax As Long
iImage As Long
iSelectedImage As Long
cChildren As Long
lParam As Long
End Type
Private Type TVITEMEX
TVI As TVITEM
iIntegral As Long
End Type
Private Type TVITEMEX_V61
TVI As TVITEM
iIntegral As Long
uStateEx As Long
hWnd As Long
iExpandedImage As Long
End Type
Private Type TVINSERTSTRUCT
hParent As Long
hInsertAfter As Long
Item As TVITEMEX
End Type
Private Type TVHITTESTINFO
PT As POINTAPI
Flags As Long
hItem As Long
End Type
Private Type TVSORTCB
hParent As Long
lpfnCompare As Long
lParam As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Const CDDS_PREPAINT As Long = &H1
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM + 1)
Private Const CDIS_SELECTED As Long = &H1
Private Const CDIS_DISABLED As Long = &H4
Private Const CDIS_FOCUS As Long = &H10
Private Const CDIS_HOT As Long = &H40
Private Const CDRF_DODEFAULT As Long = &H0
Private Const CDRF_NEWFONT As Long = &H2
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20
Private Const TVCDRF_NOIMAGES As Long = &H10000
Private Type NMCUSTOMDRAW
hdr As NMHDR
dwDrawStage As Long
hDC As Long
RC As RECT
dwItemSpec As Long
uItemState As Long
lItemlParam As Long
End Type
Private Type NMTVCUSTOMDRAW
NMCD As NMCUSTOMDRAW
ClrText As Long
ClrTextBk As Long
iLevel As Long
End Type
Private Type NMTREEVIEW
hdr As NMHDR
Action As Long
ItemOld As TVITEM
ItemNew As TVITEM
PTDrag As POINTAPI
End Type
Private Type NMTVDISPINFO
hdr As NMHDR
Item As TVITEM
End Type
Private Type NMTVGETINFOTIP
hdr As NMHDR
pszText As Long
cchTextMax As Long
hItem As Long
lParam As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event NodeClick(ByVal Node As TvwNode, ByVal Button As Integer)
Attribute NodeClick.VB_Description = "Occurs when a node is clicked."
Public Event NodeDblClick(ByVal Node As TvwNode, ByVal Button As Integer)
Attribute NodeDblClick.VB_Description = "Occurs when a node is double clicked."
Public Event NodeBeforeCheck(ByVal Node As TvwNode, ByRef Cancel As Boolean)
Attribute NodeBeforeCheck.VB_Description = "Occurs before a node is about to be checked."
Public Event NodeCheck(ByVal Node As TvwNode)
Attribute NodeCheck.VB_Description = "Occurs when a node is checked."
Public Event NodeDrag(ByVal Node As TvwNode, ByVal Button As Integer)
Attribute NodeDrag.VB_Description = "Occurs when a node initiate a drag-and-drop operation."
Public Event NodeBeforeSelect(ByVal Node As TvwNode, ByRef Cancel As Boolean)
Attribute NodeBeforeSelect.VB_Description = "Occurs before a node is about to be selected."
Public Event NodeSelect(ByVal Node As TvwNode)
Attribute NodeSelect.VB_Description = "Occurs when a node is selected."
Public Event NodeRangeSelect(ByVal Node As TvwNode, ByRef Cancel As Boolean)
Attribute NodeRangeSelect.VB_Description = "Occurs for each node when a range of nodes is about to be selected."
Public Event BeforeCollapse(ByVal Node As TvwNode, ByRef Cancel As Boolean)
Attribute BeforeCollapse.VB_Description = "Occurs before a node is about to collapse."
Public Event Collapse(ByVal Node As TvwNode)
Attribute Collapse.VB_Description = "Occurs when a node is collapsed."
Public Event BeforeExpand(ByVal Node As TvwNode, ByRef Cancel As Boolean)
Attribute BeforeExpand.VB_Description = "Occurs before a node is about to expand."
Public Event Expand(ByVal Node As TvwNode)
Attribute Expand.VB_Description = "Occurs when a node is expanded; that is, when its child nodes become visible."
Public Event BeforeLabelEdit(ByRef Cancel As Boolean)
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected node."
Public Event AfterLabelEdit(ByRef Cancel As Boolean, ByRef NewString As String)
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected node."
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
Private Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetWindowTheme Lib "uxtheme" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const ICC_TREEVIEW_CLASSES As Long = &H2
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const CF_UNICODETEXT As Long = 13
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114
Private Const SB_LINELEFT As Long = 0, SB_LINERIGHT As Long = 1
Private Const SB_LINEUP As Long = 0, SB_LINEDOWN As Long = 1
Private Const SW_HIDE As Long = &H0
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
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
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_SETREDRAW As Long = &HB
Private Const COLOR_HOTLIGHT As Long = 26
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETVERSION As Long = (CCM_FIRST + 7)
Private Const WM_USER As Long = &H400
Private Const UM_CHECKSTATECHANGED As Long = (WM_USER + 100) ' See KB 261289
Private Const UM_BUTTONDOWN As Long = (WM_USER + 500)
Private Const TVM_FIRST As Long = &H1100
Private Const TVM_INSERTITEMA As Long = (TVM_FIRST + 0)
Private Const TVM_INSERTITEMW As Long = (TVM_FIRST + 50)
Private Const TVM_INSERTITEM As Long = TVM_INSERTITEMW
Private Const TVM_DELETEITEM As Long = (TVM_FIRST + 1)
Private Const TVM_EXPAND As Long = (TVM_FIRST + 2)
Private Const TVM_GETITEMRECT As Long = (TVM_FIRST + 4)
Private Const TVM_GETCOUNT As Long = (TVM_FIRST + 5)
Private Const TVM_GETINDENT As Long = (TVM_FIRST + 6)
Private Const TVM_SETINDENT As Long = (TVM_FIRST + 7)
Private Const TVM_GETIMAGELIST As Long = (TVM_FIRST + 8)
Private Const TVM_SETIMAGELIST As Long = (TVM_FIRST + 9)
Private Const TVM_GETNEXTITEM As Long = (TVM_FIRST + 10)
Private Const TVM_SELECTITEM As Long = (TVM_FIRST + 11)
Private Const TVM_GETITEMA As Long = (TVM_FIRST + 12)
Private Const TVM_GETITEMW As Long = (TVM_FIRST + 62)
Private Const TVM_GETITEM As Long = TVM_GETITEMW
Private Const TVM_SETITEMA As Long = (TVM_FIRST + 13)
Private Const TVM_SETITEMW As Long = (TVM_FIRST + 63)
Private Const TVM_SETITEM As Long = TVM_SETITEMW
Private Const TVM_EDITLABELA As Long = (TVM_FIRST + 14)
Private Const TVM_EDITLABELW As Long = (TVM_FIRST + 65)
Private Const TVM_EDITLABEL As Long = TVM_EDITLABELW
Private Const TVM_GETEDITCONTROL As Long = (TVM_FIRST + 15)
Private Const TVM_GETVISIBLECOUNT As Long = (TVM_FIRST + 16)
Private Const TVM_HITTEST As Long = (TVM_FIRST + 17)
Private Const TVM_CREATEDRAGIMAGE As Long = (TVM_FIRST + 18)
Private Const TVM_SORTCHILDREN As Long = (TVM_FIRST + 19)
Private Const TVM_ENSUREVISIBLE As Long = (TVM_FIRST + 20)
Private Const TVM_SORTCHILDRENCB As Long = (TVM_FIRST + 21)
Private Const TVM_ENDEDITLABELNOW As Long = (TVM_FIRST + 22)
Private Const TVM_GETISEARCHSTRINGA As Long = (TVM_FIRST + 23)
Private Const TVM_GETISEARCHSTRINGW As Long = (TVM_FIRST + 64)
Private Const TVM_GETISEARCHSTRING As Long = TVM_GETISEARCHSTRINGW
Private Const TVM_SETTOOLTIPS As Long = (TVM_FIRST + 24)
Private Const TVM_GETTOOLTIPS As Long = (TVM_FIRST + 25)
Private Const TVM_SETINSERTMARK As Long = (TVM_FIRST + 26)
Private Const TVM_SETITEMHEIGHT As Long = (TVM_FIRST + 27)
Private Const TVM_GETITEMHEIGHT As Long = (TVM_FIRST + 28)
Private Const TVM_SETBKCOLOR As Long = (TVM_FIRST + 29)
Private Const TVM_SETTEXTCOLOR As Long = (TVM_FIRST + 30)
Private Const TVM_SETINSERTMARKCOLOR As Long = (TVM_FIRST + 37)
Private Const TVM_GETINSERTMARKCOLOR As Long = (TVM_FIRST + 38)
Private Const TVM_GETITEMSTATE As Long = (TVM_FIRST + 39)
Private Const TVM_SETLINECOLOR As Long = (TVM_FIRST + 40)
Private Const TVM_GETLINECOLOR As Long = (TVM_FIRST + 41)
Private Const TVM_SETEXTENDEDSTYLE As Long = (TVM_FIRST + 44)
Private Const TVM_GETEXTENDEDSTYLE As Long = (TVM_FIRST + 45)
Private Const TVN_FIRST As Long = (-400)
Private Const TVN_SELCHANGINGA As Long = (TVN_FIRST - 1)
Private Const TVN_SELCHANGINGW As Long = (TVN_FIRST - 50)
Private Const TVN_SELCHANGING As Long = TVN_SELCHANGINGW
Private Const TVN_SELCHANGEDA As Long = (TVN_FIRST - 2)
Private Const TVN_SELCHANGEDW As Long = (TVN_FIRST - 51)
Private Const TVN_SELCHANGED As Long = TVN_SELCHANGEDW
Private Const TVN_GETDISPINFOA As Long = (TVN_FIRST - 3)
Private Const TVN_GETDISPINFOW As Long = (TVN_FIRST - 52)
Private Const TVN_GETDISPINFO As Long = TVN_GETDISPINFOW
Private Const TVN_SETDISPINFOA As Long = (TVN_FIRST - 4)
Private Const TVN_SETDISPINFOW As Long = (TVN_FIRST - 53)
Private Const TVN_SETDISPINFO As Long = TVN_SETDISPINFOW
Private Const TVN_ITEMEXPANDINGA As Long = (TVN_FIRST - 5)
Private Const TVN_ITEMEXPANDINGW As Long = (TVN_FIRST - 54)
Private Const TVN_ITEMEXPANDING As Long = TVN_ITEMEXPANDINGW
Private Const TVN_ITEMEXPANDEDA As Long = (TVN_FIRST - 6)
Private Const TVN_ITEMEXPANDEDW As Long = (TVN_FIRST - 55)
Private Const TVN_ITEMEXPANDED As Long = TVN_ITEMEXPANDEDW
Private Const TVN_BEGINDRAGA As Long = (TVN_FIRST - 7)
Private Const TVN_BEGINDRAGW As Long = (TVN_FIRST - 56)
Private Const TVN_BEGINDRAG As Long = TVN_BEGINDRAGW
Private Const TVN_BEGINRDRAGA As Long = (TVN_FIRST - 8)
Private Const TVN_BEGINRDRAGW As Long = (TVN_FIRST - 57)
Private Const TVN_BEGINRDRAG As Long = TVN_BEGINRDRAGW
Private Const TVN_DELETEITEMA As Long = (TVN_FIRST - 9)
Private Const TVN_DELETEITEMW As Long = (TVN_FIRST - 58)
Private Const TVN_DELETEITEM As Long = TVN_DELETEITEMW
Private Const TVN_BEGINLABELEDITA As Long = (TVN_FIRST - 10)
Private Const TVN_BEGINLABELEDITW As Long = (TVN_FIRST - 59)
Private Const TVN_BEGINLABELEDIT As Long = TVN_BEGINLABELEDITW
Private Const TVN_ENDLABELEDITA As Long = (TVN_FIRST - 11)
Private Const TVN_ENDLABELEDITW As Long = (TVN_FIRST - 60)
Private Const TVN_ENDLABELEDIT As Long = TVN_ENDLABELEDITW
Private Const TVN_KEYDOWN As Long = (TVN_FIRST - 12)
Private Const TVN_GETINFOTIPA As Long = (TVN_FIRST - 13)
Private Const TVN_GETINFOTIPW As Long = (TVN_FIRST - 14)
Private Const TVN_GETINFOTIP As Long = TVN_GETINFOTIPW
Private Const TVN_SINGLEEXPAND As Long = (TVN_FIRST - 15)
Private Const TVSIL_NORMAL As Long = 0
Private Const TVSIL_STATE As Long = 2
Private Const TVIF_TEXT As Long = &H1
Private Const TVIF_IMAGE As Long = &H2
Private Const TVIF_PARAM As Long = &H4
Private Const TVIF_STATE As Long = &H8
Private Const TVIF_HANDLE As Long = &H10
Private Const TVIF_SELECTEDIMAGE As Long = &H20
Private Const TVIF_CHILDREN As Long = &H40
Private Const TVIF_INTEGRAL As Long = &H80
Private Const TVIF_STATEEX As Long = &H100
Private Const TVI_ROOT As Long = &HFFFF0000
Private Const TVI_FIRST As Long = &HFFFF0001
Private Const TVI_LAST As Long = &HFFFF0002
Private Const TVI_SORT As Long = &HFFFF0003
Private Const TVC_UNKNOWN As Long = &H0
Private Const TVC_BYMOUSE As Long = &H1
Private Const TVC_BYKEYBOARD As Long = &H2
Private Const TVIS_FOCUSED As Long = &H1
Private Const TVIS_SELECTED As Long = &H2
Private Const TVIS_CUT As Long = &H4
Private Const TVIS_DROPHILITED As Long = &H8
Private Const TVIS_BOLD As Long = &H10
Private Const TVIS_EXPANDED As Long = &H20
Private Const TVIS_EXPANDEDONCE As Long = &H40
Private Const TVIS_OVERLAYMASK As Long = &HF00
Private Const TVIS_STATEIMAGEMASK As Long = &HF000&
Private Const TVIS_EX_DISABLED As Long = &H2
Private Const TVHT_NOWHERE As Long = &H1
Private Const TVHT_ONITEMICON As Long = &H2
Private Const TVHT_ONITEMLABEL As Long = &H4
Private Const TVHT_ONITEMINDENT As Long = &H8
Private Const TVHT_ONITEMBUTTON As Long = &H10
Private Const TVHT_ONITEMRIGHT As Long = &H20
Private Const TVHT_ONITEMSTATEICON As Long = &H40
Private Const TVHT_ONITEM As Long = TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON
Private Const TVHT_ABOVE As Long = &H100
Private Const TVHT_BELOW As Long = &H200
Private Const TVHT_TORIGHT As Long = &H400
Private Const TVHT_TOLEFT As Long = &H800
Private Const TVE_COLLAPSE As Long = &H1
Private Const TVE_EXPAND As Long = &H2
Private Const TVE_TOGGLE As Long = &H3
Private Const TVE_EXPANDPARTIAL As Long = &H4000
Private Const TVE_COLLAPSERESET As Long = &H8000&
Private Const TVNRET_DEFAULT As Long = 0
Private Const TVNRET_SKIPOLD As Long = 1
Private Const TVNRET_SKIPNEW As Long = 2
Private Const TVGN_ROOT As Long = &H0
Private Const TVGN_NEXT As Long = &H1
Private Const TVGN_PREVIOUS As Long = &H2
Private Const TVGN_PARENT As Long = &H3
Private Const TVGN_CHILD As Long = &H4
Private Const TVGN_FIRSTVISIBLE As Long = &H5
Private Const TVGN_NEXTVISIBLE As Long = &H6
Private Const TVGN_PREVIOUSVISIBLE As Long = &H7
Private Const TVGN_DROPHILITE As Long = &H8
Private Const TVGN_CARET As Long = &H9
Private Const TVGN_LASTVISIBLE As Long = &HA
Private Const IIL_UNCHECKED As Long = 1
Private Const IIL_CHECKED As Long = 2
Private Const I_IMAGECALLBACK As Long = (-1)
Private Const NM_FIRST As Long = 0
Private Const NM_CLICK As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK As Long = (NM_FIRST - 3)
Private Const NM_RCLICK As Long = (NM_FIRST - 5)
Private Const NM_RDBLCLK As Long = (NM_FIRST - 6)
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Private Const TVS_EX_DOUBLEBUFFER As Long = &H4
Private Const TVS_HASBUTTONS As Long = &H1
Private Const TVS_HASLINES As Long = &H2
Private Const TVS_LINESATROOT As Long = &H4
Private Const TVS_EDITLABELS As Long = &H8
Private Const TVS_DISABLEDRAGDROP As Long = &H10
Private Const TVS_SHOWSELALWAYS As Long = &H20
Private Const TVS_RTLREADING As Long = &H40
Private Const TVS_NOTOOLTIPS As Long = &H80
Private Const TVS_CHECKBOXES As Long = &H100
Private Const TVS_TRACKSELECT As Long = &H200
Private Const TVS_SINGLEEXPAND As Long = &H400
Private Const TVS_INFOTIP As Long = &H800
Private Const TVS_FULLROWSELECT As Long = &H1000
Private Const TVS_NOSCROLL As Long = &H2000
Private Const TVS_NONEVENHEIGHT As Long = &H4000
Private Const TVS_NOHSCROLL As Long = &H8000&
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private TreeViewHandle As Long, TreeViewToolTipHandle As Long
Private TreeViewFontHandle As Long
Private TreeViewIMCHandle As Long
Private TreeViewCharCodeCache As Long
Private TreeViewIsClick As Boolean
Private TreeViewMouseOver As Boolean
Private TreeViewDesignMode As Boolean
Private TreeViewLabelInEdit As Boolean
Private TreeViewStartLabelEdit As Boolean
Private TreeViewButtonDown As Integer
Private TreeViewDragItemBuffer As Long, TreeViewDragItem As Long
Private TreeViewInsertMarkItem As Long, TreeViewInsertMarkAfter As Boolean
Private TreeViewExpandItem As Long, TreeViewPrevExpandItem As Long, TreeViewTickCount As Double
Private TreeViewSampleMode As Boolean
Private TreeViewImageListObjectPointer As Long
Private TreeViewAlignable As Boolean
Private TreeViewFocused As Boolean
Private TreeViewSelectedCount As Long
Private TreeViewSelectedItems() As Long
Private TreeViewClickSelectedCount As Long
Private TreeViewClickShift As Integer
Private TreeViewAnchorItem As Long
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private DispIDImageList As Long, ImageListArray() As String
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropNodes As TvwNodes
Private PropSelectedNodes As TvwSelectedNodes
Private PropVisualStyles As Boolean
Private PropVisualTheme As TvwVisualThemeConstants
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropOLEDragDropScroll As Boolean
Private PropOLEDragExpandTime As Long
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropImageListName As String, PropImageListInit As Boolean
Private PropBorderStyle As CCBorderStyleConstants
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR
Private PropRedraw As Boolean
Private PropStyle As TvwStyleConstants
Private PropLineStyle As TvwLineStyleConstants
Private PropLineColor As OLE_COLOR
Private PropLabelEdit As TvwLabelEditConstants
Private PropCheckboxes As Boolean
Private PropShowTips As Boolean
Private PropHideSelection As Boolean
Private PropFullRowSelect As Boolean
Private PropHotTracking As Boolean
Private PropIndentation As Long
Private PropPathSeparator As String
Private PropScroll As Boolean
Private PropSingleSel As Boolean
Private PropSorted As Boolean
Private PropSortOrder As TvwSortOrderConstants
Private PropSortType As TvwSortTypeConstants
Private PropInsertMarkColor As OLE_COLOR
Private PropDoubleBuffer As Boolean
Private PropIMEMode As CCIMEModeConstants
Private PropMultiSelect As TvwMultiSelectConstants

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
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd, vbKeyReturn, vbKeyEscape
            If TreeViewLabelInEdit = False Then
                If (KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape) And IsInputKey = False Then Exit Sub
            End If
            SendMessage hWnd, wMsg, wParam, ByVal lParam
            Handled = True
        Case vbKeyTab
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
Call ComCtlsInitCC(ICC_TREEVIEW_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim ImageListArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then TreeViewAlignable = False Else TreeViewAlignable = True
TreeViewDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropVisualTheme = TvwVisualThemeStandard
PropOLEDragMode = vbOLEDragManual
PropOLEDragDropScroll = True
PropOLEDragExpandTime = -1
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = "(None)"
PropBorderStyle = CCBorderStyleSunken
PropBackColor = vbWindowBackground
PropForeColor = vbWindowText
PropRedraw = True
PropStyle = TvwStyleTreeLinesPlusMinusPictureText
PropLineStyle = TvwLineStyleTreeLines
PropLineColor = vbGrayText
PropLabelEdit = TvwLabelEditAutomatic
PropCheckboxes = False
PropShowTips = False
PropHideSelection = True
PropFullRowSelect = False
PropHotTracking = False
PropIndentation = 0
PropPathSeparator = "\"
PropScroll = True
PropSingleSel = False
PropSorted = False
PropSortOrder = TvwSortOrderAscending
PropSortType = TvwSortTypeBinary
PropInsertMarkColor = vbBlack
PropDoubleBuffer = True
PropIMEMode = CCIMEModeNoControl
PropMultiSelect = TvwMultiSelectNone
Call CreateTreeView
If TreeViewDesignMode = True Then
    TreeViewSampleMode = True
    Dim SampleNode As New TvwNode
    SampleNode.FInit Me, vbNullString, 1, 1, 1, 1
    Me.FNodesAdd SampleNode, , TvwNodeRelationshipFirst, "Sample Node", 1, 1
    Me.FNodesAdd Nothing, SampleNode, TvwNodeRelationshipChild, "Sample Node", 1, 1
    Me.FNodesAdd Nothing, SampleNode, TvwNodeRelationshipChild, "Sample Node", 1, 1
    Me.FNodesAdd Nothing, , TvwNodeRelationshipNext, "Sample Node", 1, 1
    SampleNode.Expanded = True
    TreeViewSampleMode = False
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then TreeViewAlignable = False Else TreeViewAlignable = True
TreeViewDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
PropVisualTheme = .ReadProperty("VisualTheme", TvwVisualThemeStandard)
Me.Enabled = .ReadProperty("Enabled", True)
PropOLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
PropOLEDragDropScroll = .ReadProperty("OLEDragDropScroll", True)
PropOLEDragExpandTime = .ReadProperty("OLEDragExpandTime", -1)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = .ReadProperty("ImageList", "(None)")
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSunken)
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
PropForeColor = .ReadProperty("ForeColor", vbWindowText)
PropRedraw = .ReadProperty("Redraw", True)
PropStyle = .ReadProperty("Style", TvwStyleTreeLinesPlusMinusPictureText)
PropLineStyle = .ReadProperty("LineStyle", TvwLineStyleTreeLines)
PropLineColor = .ReadProperty("LineColor", vbGrayText)
PropLabelEdit = .ReadProperty("LabelEdit", TvwLabelEditAutomatic)
PropCheckboxes = .ReadProperty("Checkboxes", False)
PropShowTips = .ReadProperty("ShowTips", False)
PropHideSelection = .ReadProperty("HideSelection", True)
PropFullRowSelect = .ReadProperty("FullRowSelect", False)
PropHotTracking = .ReadProperty("HotTracking", False)
PropIndentation = (.ReadProperty("Indentation", 0) * PixelsPerDIP_X())
PropPathSeparator = VarToStr(.ReadProperty("PathSeparator", "\"))
PropScroll = .ReadProperty("Scroll", True)
PropSingleSel = .ReadProperty("SingleSel", False)
PropSorted = .ReadProperty("Sorted", False)
PropSortOrder = .ReadProperty("SortOrder", TvwSortOrderAscending)
PropSortType = .ReadProperty("SortType", TvwSortTypeBinary)
PropInsertMarkColor = .ReadProperty("InsertMarkColor", vbBlack)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
PropIMEMode = .ReadProperty("IMEMode", CCIMEModeNoControl)
PropMultiSelect = .ReadProperty("MultiSelect", TvwMultiSelectNone)
End With
Call CreateTreeView
If TreeViewDesignMode = True Then
    TreeViewSampleMode = True
    Dim SampleNode As New TvwNode
    SampleNode.FInit Me, vbNullString, 1, 1, 1, 1
    Me.FNodesAdd SampleNode, , TvwNodeRelationshipFirst, "Sample Node", 1, 1
    Me.FNodesAdd Nothing, SampleNode, TvwNodeRelationshipChild, "Sample Node", 1, 1
    Me.FNodesAdd Nothing, SampleNode, TvwNodeRelationshipChild, "Sample Node", 1, 1
    Me.FNodesAdd Nothing, , TvwNodeRelationshipNext, "Sample Node", 1, 1
    SampleNode.Expanded = True
    TreeViewSampleMode = False
End If
If Not PropImageListName = "(None)" Then TimerImageList.Enabled = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "VisualTheme", PropVisualTheme, TvwVisualThemeStandard
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDragDropScroll", PropOLEDragDropScroll, True
.WriteProperty "OLEDragExpandTime", PropOLEDragExpandTime, -1
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "ImageList", PropImageListName, "(None)"
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSunken
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "ForeColor", PropForeColor, vbWindowText
.WriteProperty "Redraw", PropRedraw, True
.WriteProperty "Style", PropStyle, TvwStyleTreeLinesPlusMinusPictureText
.WriteProperty "LineStyle", PropLineStyle, TvwLineStyleTreeLines
.WriteProperty "LineColor", PropLineColor, vbGrayText
.WriteProperty "LabelEdit", PropLabelEdit, TvwLabelEditAutomatic
.WriteProperty "Checkboxes", PropCheckboxes, False
.WriteProperty "ShowTips", PropShowTips, False
.WriteProperty "HideSelection", PropHideSelection, True
.WriteProperty "FullRowSelect", PropFullRowSelect, False
.WriteProperty "HotTracking", PropHotTracking, False
.WriteProperty "Indentation", (PropIndentation / PixelsPerDIP_X()), 0
.WriteProperty "PathSeparator", IIf(PropPathSeparator = "\", "\", StrToVar(PropPathSeparator)), "\"
.WriteProperty "Scroll", PropScroll, True
.WriteProperty "SingleSel", PropSingleSel, False
.WriteProperty "Sorted", PropSorted, False
.WriteProperty "SortOrder", PropSortOrder, TvwSortOrderAscending
.WriteProperty "SortType", PropSortType, TvwSortTypeBinary
.WriteProperty "InsertMarkColor", PropInsertMarkColor, vbBlack
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
.WriteProperty "IMEMode", PropIMEMode, CCIMEModeNoControl
.WriteProperty "MultiSelect", PropMultiSelect, TvwMultiSelectNone
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
TreeViewDragItem = 0
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
If TreeViewHandle <> 0 Then
    If State = vbOver And Not Effect = vbDropEffectNone Then
        If PropOLEDragDropScroll = True Then
            Dim RC As RECT
            GetWindowRect TreeViewHandle, RC
            Dim dwStyle As Long
            dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
            If (dwStyle And WS_HSCROLL) = WS_HSCROLL Then
                If Abs(X) < (16 * PixelsPerDIP_X()) Then
                    SendMessage TreeViewHandle, WM_HSCROLL, SB_LINELEFT, ByVal 0&
                ElseIf Abs(X - (RC.Right - RC.Left)) < (16 * PixelsPerDIP_X()) Then
                    SendMessage TreeViewHandle, WM_HSCROLL, SB_LINERIGHT, ByVal 0&
                End If
            End If
            If (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                If Abs(Y) < (16 * PixelsPerDIP_Y()) Then
                    SendMessage TreeViewHandle, WM_VSCROLL, SB_LINEUP, ByVal 0&
                ElseIf Abs(Y - (RC.Bottom - RC.Top)) < (16 * PixelsPerDIP_Y()) Then
                    SendMessage TreeViewHandle, WM_VSCROLL, SB_LINEDOWN, ByVal 0&
                End If
            End If
        End If
        Dim TVHTI As TVHITTESTINFO
        With TVHTI
        .PT.X = X
        .PT.Y = Y
        SendMessage TreeViewHandle, TVM_HITTEST, 0, ByVal VarPtr(TVHTI)
        If .hItem <> 0 And (.Flags And TVHT_ONITEMBUTTON) <> 0 Then
            If Me.FNodeExpanded(.hItem) = False And PropOLEDragExpandTime > 0 Or PropOLEDragExpandTime = -1 Then
                TreeViewExpandItem = .hItem
                If TreeViewExpandItem <> TreeViewPrevExpandItem Then
                    TreeViewTickCount = 0
                Else
                    If TreeViewTickCount = 0 Then
                        TreeViewTickCount = CLngToULng(GetTickCount())
                    ElseIf (CLngToULng(GetTickCount()) - TreeViewTickCount) > IIf(PropOLEDragExpandTime > 0, PropOLEDragExpandTime, GetDoubleClickTime() * 2) Then
                        Me.FNodeExpanded(.hItem) = True
                    End If
                End If
                TreeViewPrevExpandItem = TreeViewExpandItem
            Else
                TreeViewPrevExpandItem = 0
            End If
        Else
            TreeViewPrevExpandItem = 0
        End If
        End With
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
If TreeViewDragItem <> 0 Then
    If PropOLEDragMode = vbOLEDragAutomatic Then
        Dim Text As String
        Text = Me.FNodeText(TreeViewDragItem)
        Data.SetData StrToVar(Text & vbNullChar), CF_UNICODETEXT
        Data.SetData Text, vbCFText
        Const vbDropEffectLink As Long = 4 ' Undocumented
        AllowedEffects = vbDropEffectCopy Or vbDropEffectMove Or vbDropEffectLink
    End If
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then TreeViewDragItem = 0
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
If TreeViewDragItem <> 0 Then Exit Sub
If TreeViewDragItemBuffer <> 0 Then TreeViewDragItem = TreeViewDragItemBuffer
UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
Static LastHeight As Single, LastWidth As Single, LastAlign As Integer
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl.Extender
Dim Align As Integer
If TreeViewAlignable = True Then Align = .Align Else Align = vbAlignNone
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
If TreeViewHandle <> 0 Then MoveWindow TreeViewHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyTreeView
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropImageListInit = False Then
    If Not PropImageListName = "(None)" Then Me.ImageList = PropImageListName
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
hWnd = TreeViewHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndLabelEdit() As Long
Attribute hWndLabelEdit.VB_Description = "Returns a handle to a control."
If TreeViewHandle <> 0 Then hWndLabelEdit = SendMessage(TreeViewHandle, TVM_GETEDITCONTROL, 0, ByVal 0&)
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
OldFontHandle = TreeViewFontHandle
TreeViewFontHandle = CreateGDIFontFromOLEFont(PropFont)
If TreeViewHandle <> 0 Then SendMessage TreeViewHandle, WM_SETFONT, TreeViewFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = TreeViewFontHandle
TreeViewFontHandle = CreateGDIFontFromOLEFont(PropFont)
If TreeViewHandle <> 0 Then SendMessage TreeViewHandle, WM_SETFONT, TreeViewFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If TreeViewHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        If PropVisualTheme = TvwVisualThemeExplorer Then
            SetWindowTheme TreeViewHandle, StrPtr("Explorer"), 0
        Else
            ActivateVisualStyles TreeViewHandle
        End If
    Else
        RemoveVisualStyles TreeViewHandle
    End If
    Call SetVisualStylesToolTip
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get VisualTheme() As TvwVisualThemeConstants
Attribute VisualTheme.VB_Description = "Returns/sets the visual theme. Requires comctl32.dll version 6.0 or higher."
VisualTheme = PropVisualTheme
End Property

Public Property Let VisualTheme(ByVal Value As TvwVisualThemeConstants)
Select Case Value
    Case TvwVisualThemeStandard, TvwVisualThemeExplorer
        PropVisualTheme = Value
    Case Else
        Err.Raise 380
End Select
Me.VisualStyles = PropVisualStyles
UserControl.PropertyChanged "VisualTheme"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If TreeViewHandle <> 0 Then EnableWindow TreeViewHandle, IIf(Value = True, 1, 0)
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

Public Property Get OLEDragExpandTime() As Long
Attribute OLEDragExpandTime.VB_Description = "Returns/sets the OLE drag expand time in milliseconds. A value of 0 indicates that auto-expansion is disabled. A value of -1 indicates that the system's double click time, multiplied with 2, is used."
OLEDragExpandTime = PropOLEDragExpandTime
End Property

Public Property Let OLEDragExpandTime(ByVal Value As Long)
If Value < -1 Then
    If TreeViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropOLEDragExpandTime = Value
UserControl.PropertyChanged "OLEDragExpandTime"
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
If TreeViewDesignMode = False Then Call RefreshMousePointer
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
        If TreeViewDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If TreeViewDesignMode = False Then Call RefreshMousePointer
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
If TreeViewDesignMode = False Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
    dwMask = 0
End If
If PropRightToLeft = True Then
    If PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
End If
If TreeViewHandle <> 0 Then
    Call ComCtlsSetRightToLeft(TreeViewHandle, dwMask)
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If (dwStyle And TVS_RTLREADING) = TVS_RTLREADING Then dwStyle = dwStyle And Not TVS_RTLREADING
    If PropRightToLeft = True And PropRightToLeftLayout = False Then dwStyle = dwStyle Or TVS_RTLREADING
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
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
If TreeViewDesignMode = False Then
    If PropImageListInit = False And TreeViewImageListObjectPointer = 0 Then
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
If TreeViewHandle <> 0 Then
    Dim Success As Boolean, Handle As Long
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> 0)
        End If
        If Success = True Then
            Select Case PropStyle
                Case TvwStylePictureText, TvwStylePlusMinusPictureText, TvwStyleTreeLinesPictureText, TvwStyleTreeLinesPlusMinusPictureText
                    SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal Handle
                Case Else
                    SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal 0&
            End Select
            TreeViewImageListObjectPointer = ObjPtr(Value)
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
                        Select Case PropStyle
                            Case TvwStylePictureText, TvwStylePlusMinusPictureText, TvwStyleTreeLinesPictureText, TvwStyleTreeLinesPlusMinusPictureText
                                SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal Handle
                            Case Else
                                SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal 0&
                        End Select
                        If TreeViewDesignMode = False Then TreeViewImageListObjectPointer = ObjPtr(ControlEnum)
                        PropImageListName = Value
                        Exit For
                    ElseIf TreeViewDesignMode = True Then
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
        SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal 0&
        TreeViewImageListObjectPointer = 0
        PropImageListName = "(None)"
    ElseIf Handle = 0 Then
        SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal 0&
    End If
    SendMessage TreeViewHandle, TVM_SETINDENT, PropIndentation, ByVal 0&
    If PropCheckboxes = True Then
        Me.Checkboxes = False
        Me.Checkboxes = True
    End If
End If
UserControl.PropertyChanged "ImageList"
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
If TreeViewHandle <> 0 Then Call ComCtlsChangeBorderStyle(TreeViewHandle, PropBorderStyle)
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If TreeViewHandle <> 0 Then
    SendMessage TreeViewHandle, TVM_SETBKCOLOR, 0, ByVal WinColor(PropBackColor)
    Me.Refresh
End If
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If TreeViewHandle <> 0 Then
    SendMessage TreeViewHandle, TVM_SETTEXTCOLOR, 0, ByVal WinColor(PropForeColor)
    Me.Refresh
End If
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Returns/sets a value that determines whether or not the tree view redraws when changing the nodes. You can speed up the creation of large lists by disabling this property before adding the nodes."
Redraw = PropRedraw
End Property

Public Property Let Redraw(ByVal Value As Boolean)
PropRedraw = Value
If TreeViewHandle <> 0 And TreeViewDesignMode = False Then
    SendMessage TreeViewHandle, WM_SETREDRAW, IIf(PropRedraw = True, 1, 0), ByVal 0&
    If PropRedraw = True Then Me.Refresh
End If
UserControl.PropertyChanged "Redraw"
End Property

Public Property Get Style() As TvwStyleConstants
Attribute Style.VB_Description = "Returns/sets the style."
Style = PropStyle
End Property

Public Property Let Style(ByVal Value As TvwStyleConstants)
Select Case Value
    Case TvwStyleTextOnly, TvwStylePictureText, TvwStylePlusMinusText, TvwStylePlusMinusPictureText, TvwStyleTreeLinesText, TvwStyleTreeLinesPictureText, TvwStyleTreeLinesPlusMinusText, TvwStyleTreeLinesPlusMinusPictureText
        PropStyle = Value
    Case Else
        Err.Raise 380
End Select
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If (dwStyle And TVS_HASBUTTONS) = TVS_HASBUTTONS Then dwStyle = dwStyle And Not TVS_HASBUTTONS
    If (dwStyle And TVS_HASLINES) = TVS_HASLINES Then dwStyle = dwStyle And Not TVS_HASLINES
    Select Case PropStyle
        Case TvwStylePlusMinusText, TvwStylePlusMinusPictureText
            dwStyle = dwStyle Or TVS_HASBUTTONS
        Case TvwStyleTreeLinesText, TvwStyleTreeLinesPictureText
            dwStyle = dwStyle Or TVS_HASLINES
        Case TvwStyleTreeLinesPlusMinusText, TvwStyleTreeLinesPlusMinusPictureText
            dwStyle = dwStyle Or TVS_HASLINES Or TVS_HASBUTTONS
    End Select
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
    If SendMessage(TreeViewHandle, TVM_GETIMAGELIST, TVSIL_NORMAL, ByVal 0&) <> 0 Then
        Select Case PropStyle
            Case TvwStyleTextOnly, TvwStylePlusMinusText, TvwStyleTreeLinesText, TvwStyleTreeLinesPlusMinusText
                SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal 0&
        End Select
    Else
        Select Case PropStyle
            Case TvwStylePictureText, TvwStylePlusMinusPictureText, TvwStyleTreeLinesPictureText, TvwStyleTreeLinesPlusMinusPictureText
                If Not PropImageListControl Is Nothing Then
                    SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal CLng(PropImageListControl.hImageList)
                ElseIf Not PropImageListName = "(None)" Then
                    Me.ImageList = PropImageListName
                End If
        End Select
    End If
End If
UserControl.PropertyChanged "Style"
End Property

Public Property Get LineStyle() As TvwLineStyleConstants
Attribute LineStyle.VB_Description = "Returns/sets the style of lines displayed between nodes."
LineStyle = PropLineStyle
End Property

Public Property Let LineStyle(ByVal Value As TvwLineStyleConstants)
Select Case Value
    Case TvwLineStyleTreeLines, TvwLineStyleRootLines
        PropLineStyle = Value
    Case Else
        Err.Raise 380
End Select
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    Select Case PropLineStyle
        Case TvwLineStyleTreeLines
            If (dwStyle And TVS_LINESATROOT) = TVS_LINESATROOT Then dwStyle = dwStyle And Not TVS_LINESATROOT
        Case TvwLineStyleRootLines
            If Not (dwStyle And TVS_LINESATROOT) = TVS_LINESATROOT Then dwStyle = dwStyle Or TVS_LINESATROOT
    End Select
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "LineStyle"
End Property

Public Property Get LineColor() As OLE_COLOR
Attribute LineColor.VB_Description = "Returns/sets the line color."
LineColor = PropLineColor
End Property

Public Property Let LineColor(ByVal Value As OLE_COLOR)
PropLineColor = Value
If TreeViewHandle <> 0 Then SendMessage TreeViewHandle, TVM_SETLINECOLOR, 0, ByVal WinColor(PropLineColor)
UserControl.PropertyChanged "LineColor"
End Property

Public Property Get LabelEdit() As TvwLabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a node."
LabelEdit = PropLabelEdit
End Property

Public Property Let LabelEdit(ByVal Value As TvwLabelEditConstants)
Select Case Value
    Case TvwLabelEditAutomatic, TvwLabelEditManual, TvwLabelEditDisabled
        PropLabelEdit = Value
    Case Else
        Err.Raise 380
End Select
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    Select Case PropLabelEdit
        Case TvwLabelEditAutomatic, TvwLabelEditManual
            If Not (dwStyle And TVS_EDITLABELS) = TVS_EDITLABELS Then dwStyle = dwStyle Or TVS_EDITLABELS
        Case TvwLabelEditDisabled
            If (dwStyle And TVS_EDITLABELS) = TVS_EDITLABELS Then dwStyle = dwStyle And Not TVS_EDITLABELS
    End Select
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "LabelEdit"
End Property

Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value which determines if the control displays a checkbox next to each item in the tree."
Checkboxes = PropCheckboxes
End Property

Public Property Let Checkboxes(ByVal Value As Boolean)
PropCheckboxes = Value
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If PropCheckboxes = True Then
        If Not (dwStyle And TVS_CHECKBOXES) = TVS_CHECKBOXES Then SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle Or TVS_CHECKBOXES
    Else
        If (dwStyle And TVS_CHECKBOXES) = TVS_CHECKBOXES Then
            SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle And Not TVS_CHECKBOXES
            Dim hImageList As Long
            hImageList = SendMessage(TreeViewHandle, TVM_GETIMAGELIST, TVSIL_STATE, ByVal 0&)
            If hImageList <> 0 Then
                SendMessage TreeViewHandle, TVM_SETIMAGELIST, TVSIL_STATE, ByVal 0&
                ImageList_Destroy hImageList
            End If
        End If
    End If
End If
UserControl.PropertyChanged "Checkboxes"
End Property

Public Property Get ShowTips() As Boolean
Attribute ShowTips.VB_Description = "Returns/sets a value that determines whether the tool tip text properties will be displayed or not."
ShowTips = PropShowTips
End Property

Public Property Let ShowTips(ByVal Value As Boolean)
PropShowTips = Value
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If PropShowTips = True Then
        If Not (dwStyle And TVS_INFOTIP) = TVS_INFOTIP Then dwStyle = dwStyle Or TVS_INFOTIP
    Else
        If (dwStyle And TVS_INFOTIP) = TVS_INFOTIP Then dwStyle = dwStyle And Not TVS_INFOTIP
    End If
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "ShowTips"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value that determines whether the selected item will display as selected when the tree view loses focus or not."
HideSelection = PropHideSelection
End Property

Public Property Let HideSelection(ByVal Value As Boolean)
PropHideSelection = Value
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If PropHideSelection = True Then
        If (dwStyle And TVS_SHOWSELALWAYS) = TVS_SHOWSELALWAYS Then dwStyle = dwStyle And Not TVS_SHOWSELALWAYS
    Else
        If Not (dwStyle And TVS_SHOWSELALWAYS) = TVS_SHOWSELALWAYS Then dwStyle = dwStyle Or TVS_SHOWSELALWAYS
    End If
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "HideSelection"
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets whether selecting a node highlights the entire row."
FullRowSelect = PropFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal Value As Boolean)
PropFullRowSelect = Value
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If PropFullRowSelect = True Then
        If Not (dwStyle And TVS_FULLROWSELECT) = TVS_FULLROWSELECT Then dwStyle = dwStyle Or TVS_FULLROWSELECT
    Else
        If (dwStyle And TVS_FULLROWSELECT) = TVS_FULLROWSELECT Then dwStyle = dwStyle And Not TVS_FULLROWSELECT
    End If
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "FullRowSelect"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets a value which determines if items are highlighted as the mousepointer passes over them."
HotTracking = PropHotTracking
End Property

Public Property Let HotTracking(ByVal Value As Boolean)
PropHotTracking = Value
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If PropHotTracking = True Then
        If Not (dwStyle And TVS_TRACKSELECT) = TVS_TRACKSELECT Then dwStyle = dwStyle Or TVS_TRACKSELECT
    Else
        If (dwStyle And TVS_TRACKSELECT) = TVS_TRACKSELECT Then dwStyle = dwStyle And Not TVS_TRACKSELECT
    End If
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "HotTracking"
End Property

Public Property Get Indentation() As Single
Attribute Indentation.VB_Description = "Returns/sets the width of the indentation."
Indentation = UserControl.ScaleX(PropIndentation, vbPixels, vbContainerSize)
End Property

Public Property Let Indentation(ByVal Value As Single)
If Value < 0 Then
    If TreeViewDesignMode = True Then
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
If LngValue < 0 Or ErrValue <> 0 Then
    If TreeViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
Else
    PropIndentation = LngValue
    If TreeViewHandle <> 0 Then SendMessage TreeViewHandle, TVM_SETINDENT, PropIndentation, ByVal 0&
End If
UserControl.PropertyChanged "Indentation"
End Property

Public Property Get PathSeparator() As String
Attribute PathSeparator.VB_Description = "Returns/sets the delimiter string used for the path returned by the full path property."
PathSeparator = PropPathSeparator
End Property

Public Property Let PathSeparator(ByVal Value As String)
PropPathSeparator = Value
UserControl.PropertyChanged "PathSeparator"
End Property

Public Property Get Scroll() As Boolean
Attribute Scroll.VB_Description = "Returns/sets a value which determines if the tree view displays scrollbars and allows scrolling (vertical and horizontal)."
Scroll = PropScroll
End Property

Public Property Let Scroll(ByVal Value As Boolean)
PropScroll = Value
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If PropScroll = True Then
        If (dwStyle And TVS_NOSCROLL) = TVS_NOSCROLL Then dwStyle = dwStyle And Not TVS_NOSCROLL
    Else
        If Not (dwStyle And TVS_NOSCROLL) = TVS_NOSCROLL Then dwStyle = dwStyle Or TVS_NOSCROLL
    End If
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "Scroll"
End Property

Public Property Get SingleSel() As Boolean
Attribute SingleSel.VB_Description = "Returns/sets a value which determines if selecting a new item in the tree expands that item and collapses the previously selected item."
SingleSel = PropSingleSel
End Property

Public Property Let SingleSel(ByVal Value As Boolean)
PropSingleSel = Value
If TreeViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(TreeViewHandle, GWL_STYLE)
    If PropSingleSel = True Then
        If Not (dwStyle And TVS_SINGLEEXPAND) = TVS_SINGLEEXPAND Then dwStyle = dwStyle Or TVS_SINGLEEXPAND
    Else
        If (dwStyle And TVS_SINGLEEXPAND) = TVS_SINGLEEXPAND Then dwStyle = dwStyle And Not TVS_SINGLEEXPAND
    End If
    SetWindowLong TreeViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "SingleSel"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Returns/sets a value indicating if the nodes are automatically sorted."
Sorted = PropSorted
End Property

Public Property Let Sorted(ByVal Value As Boolean)
PropSorted = Value
If PropSorted = True And TreeViewDesignMode = False Then Call SortNodes(TVI_ROOT, PropSortType)
UserControl.PropertyChanged "Sorted"
End Property

Public Property Get SortOrder() As TvwSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets a value that determines whether the nodes will be sorted in ascending or descending order."
SortOrder = PropSortOrder
End Property

Public Property Let SortOrder(ByVal Value As TvwSortOrderConstants)
Select Case Value
    Case TvwSortOrderAscending, TvwSortOrderDescending
        PropSortOrder = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True And TreeViewDesignMode = False Then Call SortNodes(TVI_ROOT, PropSortType)
End Property

Public Property Get SortType() As TvwSortTypeConstants
Attribute SortType.VB_Description = "Returns/sets the sort type."
SortType = PropSortType
End Property

Public Property Let SortType(ByVal Value As TvwSortTypeConstants)
Select Case Value
    Case TvwSortTypeBinary, TvwSortTypeText
        PropSortType = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True And TreeViewDesignMode = False Then Call SortNodes(TVI_ROOT, PropSortType)
End Property

Public Property Get InsertMarkColor() As OLE_COLOR
Attribute InsertMarkColor.VB_Description = "Returns/sets the color of the insertion mark."
InsertMarkColor = PropInsertMarkColor
End Property

Public Property Let InsertMarkColor(ByVal Value As OLE_COLOR)
PropInsertMarkColor = Value
If TreeViewHandle <> 0 Then SendMessage TreeViewHandle, TVM_SETINSERTMARKCOLOR, 0, ByVal WinColor(PropInsertMarkColor)
UserControl.PropertyChanged "InsertMarkColor"
End Property

Public Property Get DoubleBuffer() As Boolean
Attribute DoubleBuffer.VB_Description = "Returns/sets a value that determines whether the control paints via double-buffering, which reduces flicker. Requires comctl32.dll version 6.1 or higher."
DoubleBuffer = PropDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal Value As Boolean)
PropDoubleBuffer = Value
If TreeViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If PropDoubleBuffer = True Then
        SendMessage TreeViewHandle, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, ByVal TVS_EX_DOUBLEBUFFER
    Else
        SendMessage TreeViewHandle, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, ByVal 0&
    End If
End If
UserControl.PropertyChanged "DoubleBuffer"
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
If TreeViewHandle <> 0 And TreeViewDesignMode = False Then
    If GetFocus() = TreeViewHandle Then
        Call ComCtlsSetIMEMode(TreeViewHandle, TreeViewIMCHandle, PropIMEMode)
    ElseIf TreeViewLabelInEdit = True Then
        Dim LabelEditHandle As Long
        LabelEditHandle = Me.hWndLabelEdit
        If LabelEditHandle <> 0 Then
            If GetFocus() = LabelEditHandle Then Call ComCtlsSetIMEMode(LabelEditHandle, TreeViewIMCHandle, PropIMEMode)
        End If
    End If
End If
UserControl.PropertyChanged "IMEMode"
End Property

Public Property Get MultiSelect() As TvwMultiSelectConstants
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the tree view and how the multiple selections can be made."
MultiSelect = PropMultiSelect
End Property

Public Property Let MultiSelect(ByVal Value As TvwMultiSelectConstants)
Select Case Value
    Case TvwMultiSelectNone, TvwMultiSelectAll, TvwMultiSelectVisibleOnly, TvwMultiSelectRestrictSiblings
        PropMultiSelect = Value
    Case Else
        Err.Raise 380
End Select
TreeViewAnchorItem = ClearSelectedItems()
TreeViewClickSelectedCount = 0
TreeViewClickShift = 0
If PropMultiSelect = TvwMultiSelectNone Then Set PropSelectedNodes = Nothing
UserControl.PropertyChanged "MultiSelect"
End Property

Public Property Get Nodes() As TvwNodes
Attribute Nodes.VB_Description = "Returns a reference to a collection of the node objects."
If PropNodes Is Nothing Then
    Set PropNodes = New TvwNodes
    PropNodes.FInit Me
End If
Set Nodes = PropNodes
End Property

Friend Sub FNodesAdd(ByRef NewNode As TvwNode, Optional ByVal RelativeNode As TvwNode, Optional ByVal Relationship As TvwNodeRelationshipConstants, Optional ByVal Text As String, Optional ByVal ImageIndex As Long, Optional ByVal SelectedImageIndex As Long)
If TreeViewHandle <> 0 Then
    Dim TVIS As TVINSERTSTRUCT, hRelative As Long, hNode As Long
    With TVIS
    If RelativeNode Is Nothing Then
        hRelative = TVI_ROOT
    Else
        hRelative = RelativeNode.Handle
    End If
    Select Case Relationship
        Case TvwNodeRelationshipFirst
            If Not hRelative = TVI_ROOT Then hRelative = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
            .hParent = hRelative
            .hInsertAfter = TVI_FIRST
        Case TvwNodeRelationshipLast
            If Not hRelative = TVI_ROOT Then hRelative = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
            .hParent = hRelative
            .hInsertAfter = TVI_LAST
        Case TvwNodeRelationshipNext
            If hRelative = TVI_ROOT Then
                .hParent = hRelative
            Else
                .hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
                If .hParent = 0 Then .hParent = TVI_ROOT
            End If
            .hInsertAfter = hRelative
        Case TvwNodeRelationshipPrevious
            Dim hPrevious As Long
            hPrevious = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PREVIOUS, ByVal hRelative)
            If hPrevious = 0 Then
                .hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
                .hInsertAfter = TVI_FIRST
            Else
                .hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
                .hInsertAfter = hPrevious
            End If
            If .hParent = 0 Then .hParent = TVI_ROOT
        Case TvwNodeRelationshipChild
            .hParent = hRelative
        Case Else
            Err.Raise 380
    End Select
    With .Item
    With .TVI
    .Mask = TVIF_TEXT Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM Or TVIF_INTEGRAL
    .pszText = StrPtr(Text)
    .cchTextMax = Len(Text) + 1
    If TreeViewSampleMode = False Then
        .iImage = I_IMAGECALLBACK
        .iSelectedImage = I_IMAGECALLBACK
    Else
        .iImage = ImageIndex - 1
        .iSelectedImage = SelectedImageIndex - 1
    End If
    .lParam = ObjPtr(NewNode)
    End With
    .iIntegral = 1
    End With
    hNode = SendMessage(TreeViewHandle, TVM_INSERTITEM, 0, ByVal VarPtr(TVIS))
    If .Item.TVI.lParam <> 0 Then
        NewNode.Handle = hNode
        If .hParent = TVI_ROOT Then
            If PropSorted = True Then Call SortNodes(TVI_ROOT, PropSortType)
        ElseIf .hParent <> 0 Then
            If RelativeNode.Sorted = True Then Call SortNodes(.hParent, RelativeNode.SortType)
        End If
    End If
    End With
End If
End Sub

Friend Function FNodesRemove(ByVal Handle As Long) As Collection
Set FNodesRemove = New Collection
If TreeViewHandle <> 0 Then
    Call NodesRemoveRecursion(FNodesRemove, Handle)
    SendMessage TreeViewHandle, TVM_DELETEITEM, 0, ByVal Handle
End If
End Function

Friend Sub FNodesClear()
If TreeViewHandle <> 0 Then
    ' The tree view control will delete all items one by one.
    ' Ensure no caret item is set to have no unnecessary TVN_SELCHANGING/TVN_SELCHANGED notifications.
    SendMessage TreeViewHandle, TVM_SELECTITEM, TVGN_CARET, ByVal 0&
    SendMessage TreeViewHandle, TVM_DELETEITEM, 0, ByVal TVI_ROOT
End If
End Sub

Friend Property Get FNodeText(ByVal Handle As Long) As String
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_TEXT
    .hItem = Handle
    Dim Buffer As String
    Buffer = String(260, vbNullChar)
    .pszText = StrPtr(Buffer)
    .cchTextMax = 260
    End With
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    FNodeText = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Friend Property Let FNodeText(ByVal Handle As Long, ByVal Value As String)
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_TEXT
    .hItem = Handle
    .pszText = StrPtr(Value)
    .cchTextMax = Len(Value) + 1
    End With
    SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
End If
End Property

Friend Sub FNodeRedraw(ByVal Handle As Long)
If TreeViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = Handle
    If SendMessage(TreeViewHandle, TVM_GETITEMRECT, 0, ByVal VarPtr(RC)) <> 0 Then
        InvalidateRect TreeViewHandle, RC, 1
        UpdateWindow TreeViewHandle
    End If
End If
End Sub

Friend Property Get FNodeSelected(ByVal Handle As Long) As Boolean
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_SELECTED
    If PropMultiSelect = TvwMultiSelectNone Then
        SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
        FNodeSelected = CBool((.State And TVIS_SELECTED) = TVIS_SELECTED)
    Else
        If SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&) = Handle Then
            SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
            FNodeSelected = CBool((.State And TVIS_SELECTED) = TVIS_SELECTED)
        Else
            FNodeSelected = IsItemSelected(Handle)
        End If
    End If
    End With
End If
End Property

Friend Property Let FNodeSelected(ByVal Handle As Long, ByVal Value As Boolean)
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_SELECTED
    If Value = True Then
        If PropMultiSelect = TvwMultiSelectNone Then
            If SendMessage(TreeViewHandle, TVM_SELECTITEM, TVGN_CARET, ByVal Handle) <> 0 Then
                .State = TVIS_SELECTED
                SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
            End If
        Else
            Select Case SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
                Case Handle
                    .State = TVIS_SELECTED
                    SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
                Case 0
                    If SendMessage(TreeViewHandle, TVM_SELECTITEM, TVGN_CARET, ByVal Handle) <> 0 Then
                        .State = TVIS_SELECTED
                        SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
                    End If
            End Select
            Call SetItemSelected(Handle, True)
        End If
    Else
        .State = 0
        SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
        If PropMultiSelect <> TvwMultiSelectNone Then Call SetItemSelected(Handle, False)
    End If
    End With
End If
End Property

Friend Property Get FNodeCheckBox(ByVal Handle As Long) As Boolean
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_STATEIMAGEMASK
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    FNodeCheckBox = CBool(StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) <> 0)
    End With
End If
End Property

Friend Property Let FNodeCheckBox(ByVal Handle As Long, ByVal Value As Boolean)
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_STATEIMAGEMASK
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    If Value = True Then
        If StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) = 0 Then .State = IndexToStateImageMask(IIL_UNCHECKED)
    Else
        If StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) <> 0 Then .State = IndexToStateImageMask(0)
    End If
    SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
    End With
End If
End Property

Friend Property Get FNodeChecked(ByVal Handle As Long) As Boolean
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_STATEIMAGEMASK
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    FNodeChecked = CBool(StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) = IIL_CHECKED)
    End With
End If
End Property

Friend Property Let FNodeChecked(ByVal Handle As Long, ByVal Value As Boolean)
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_STATEIMAGEMASK
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    If StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) <> 0 Then
        If CBool(StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) = IIL_CHECKED) <> Value Then
            Dim Ptr As Long, Node As TvwNode, Cancel As Boolean
            Ptr = GetItemPtr(Handle)
            If Ptr <> 0 Then Set Node = PtrToObj(Ptr)
            RaiseEvent NodeBeforeCheck(Node, Cancel)
            If Cancel = False Then
                If Value = True Then
                    If StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) = IIL_UNCHECKED Then .State = IndexToStateImageMask(IIL_CHECKED)
                Else
                    If StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) = IIL_CHECKED Then .State = IndexToStateImageMask(IIL_UNCHECKED)
                End If
                SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
                RaiseEvent NodeCheck(Node)
            End If
        End If
    End If
    End With
End If
End Property

Friend Property Get FNodeBold(ByVal Handle As Long) As Boolean
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_BOLD
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    FNodeBold = CBool((.State And TVIS_BOLD) = TVIS_BOLD)
    End With
End If
End Property

Friend Property Let FNodeBold(ByVal Handle As Long, ByVal Value As Boolean)
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_BOLD
    If Value = True Then
        .State = TVIS_BOLD
    Else
        .State = 0
    End If
    SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
    End With
End If
End Property

Friend Property Get FNodeGhosted(ByVal Handle As Long) As Boolean
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_CUT
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    FNodeGhosted = CBool((.State And TVIS_CUT) = TVIS_CUT)
    End With
End If
End Property

Friend Property Let FNodeGhosted(ByVal Handle As Long, ByVal Value As Boolean)
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_CUT
    If Value = True Then
        .State = TVIS_CUT
    Else
        .State = 0
    End If
    SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
    End With
End If
End Property

Friend Property Get FNodeExpanded(ByVal Handle As Long) As Boolean
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_EXPANDED
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    FNodeExpanded = CBool((.State And TVIS_EXPANDED) = TVIS_EXPANDED)
    End With
End If
End Property

Friend Property Let FNodeExpanded(ByVal Handle As Long, ByVal Value As Boolean)
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_STATE
    .hItem = Handle
    .StateMask = TVIS_EXPANDED Or TVIS_EXPANDEDONCE
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    If CBool((.State And TVIS_EXPANDED) = TVIS_EXPANDED) <> Value Then
        Dim Ptr As Long, Node As TvwNode, Cancel As Boolean
        Ptr = GetItemPtr(Handle)
        If Ptr <> 0 Then Set Node = PtrToObj(Ptr)
        If (.State And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE Then
            ' The TVN_ITEMEXPANDING and TVN_ITEMEXPANDED notification codes are not generated in this case.
            If SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CHILD, ByVal Handle) <> 0 Then
                If Value = True Then
                    RaiseEvent BeforeExpand(Node, Cancel)
                Else
                    RaiseEvent BeforeCollapse(Node, Cancel)
                End If
            End If
        End If
        If Cancel = False Then
            If Value = True Then
                If SendMessage(TreeViewHandle, TVM_EXPAND, TVE_EXPAND, ByVal Handle) = 0 Then
                    ' Node has no child items. Get expanded once child is added.
                    If (.State And TVIS_EXPANDED) = 0 Then
                        .State = .State Or TVIS_EXPANDED
                        SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
                        Cancel = True
                    End If
                End If
            Else
                If SendMessage(TreeViewHandle, TVM_EXPAND, TVE_COLLAPSE, ByVal Handle) = 0 Then
                    ' Node has no child items. Reset expanded state.
                    If (.State And TVIS_EXPANDED) = TVIS_EXPANDED Then
                        .State = .State And Not TVIS_EXPANDED
                        SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
                        Cancel = True
                    End If
                End If
            End If
            If (.State And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE Then
                ' The TVN_ITEMEXPANDING and TVN_ITEMEXPANDED notification codes are not generated in this case.
                If Cancel = False Then
                    If Value = True Then
                        RaiseEvent Expand(Node)
                    Else
                        RaiseEvent Collapse(Node)
                    End If
                End If
            End If
        End If
    End If
    End With
End If
End Property

Friend Property Get FNodeEnabled(ByVal Handle As Long) As Boolean
If TreeViewHandle <> 0 Then
    If ComCtlsSupportLevel() >= 2 Then
        Dim TVI_V61 As TVITEMEX_V61
        With TVI_V61
        .TVI.Mask = TVIF_HANDLE Or TVIF_STATEEX
        .TVI.hItem = Handle
        SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI_V61)
        FNodeEnabled = CBool((.uStateEx And TVIS_EX_DISABLED) = 0)
        End With
    Else
        FNodeEnabled = True
    End If
End If
End Property

Friend Property Let FNodeEnabled(ByVal Handle As Long, ByVal Value As Boolean)
If TreeViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim TVI_V61 As TVITEMEX_V61
    With TVI_V61
    .TVI.Mask = TVIF_HANDLE Or TVIF_STATEEX
    .TVI.hItem = Handle
    If Value = True Then
        If (.uStateEx And TVIS_EX_DISABLED) = TVIS_EX_DISABLED Then .uStateEx = .uStateEx And Not TVIS_EX_DISABLED
    Else
        If (.uStateEx And TVIS_EX_DISABLED) = 0 Then .uStateEx = .uStateEx Or TVIS_EX_DISABLED
    End If
    SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI_V61)
    End With
End If
End Property

Friend Sub FNodeSort(ByVal Handle As Long, ByVal SortType As TvwSortTypeConstants)
Call SortNodes(Handle, SortType)
End Sub

Friend Property Get FNodeChildren(ByVal Handle As Long) As Long
If TreeViewHandle <> 0 Then
    Dim hItem As Long, Count As Long
    hItem = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CHILD, ByVal Handle)
    Do While hItem <> 0
        Count = Count + 1
        hItem = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    Loop
    FNodeChildren = Count
End If
End Property

Friend Property Get FNodeChild(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CHILD, ByVal Handle))
    If Ptr <> 0 Then Set FNodeChild = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodeLevel(ByVal Handle As Long) As Long
If TreeViewHandle <> 0 Then
    Dim hItem As Long, Count As Long
    hItem = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal Handle)
    Do While hItem <> 0
        Count = Count + 1
        hItem = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem)
    Loop
    FNodeLevel = Count
End If
End Property

Friend Property Get FNodeParent(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal Handle))
    If Ptr <> 0 Then Set FNodeParent = PtrToObj(Ptr)
End If
End Property

Friend Property Let FNodeParent(ByVal Handle As Long, Value As TvwNode)
Set Me.FNodeParent(Handle) = Value
End Property

Friend Property Set FNodeParent(ByVal Handle As Long, Value As TvwNode)
If TreeViewHandle <> 0 Then
    If Value Is Nothing Then
        Err.Raise Number:=35610, Description:="Invalid object"
    Else
        Dim hParentTest As Long
        hParentTest = Value.Handle
        Do While hParentTest <> 0
            If hParentTest = Handle Then Err.Raise Number:=35614, Description:="This would introduce a cycle"
            hParentTest = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hParentTest)
        Loop
        Dim Node As TvwNode
        Set Node = Me.SelectedItem
        MoveNodes Handle, Value.Handle, TVI_FIRST
        If Not Node Is Nothing Then Set Me.SelectedItem = Node
        SendMessage TreeViewHandle, TVM_DELETEITEM, 0, ByVal Handle
    End If
End If
End Property

Friend Property Get FNodeRoot(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_ROOT, ByVal Handle))
    If Ptr <> 0 Then Set FNodeRoot = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodeNextSibling(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal Handle))
    If Ptr <> 0 Then Set FNodeNextSibling = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodePreviousSibling(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PREVIOUS, ByVal Handle))
    If Ptr <> 0 Then Set FNodePreviousSibling = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodeFirstSibling(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long, hItem As Long, hItemBuffer As Long
    hItemBuffer = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PREVIOUS, ByVal Handle)
    Do While hItemBuffer <> 0
        hItem = hItemBuffer
        hItemBuffer = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PREVIOUS, ByVal hItem)
    Loop
    If hItem = 0 Then hItem = Handle
    Ptr = GetItemPtr(hItem)
    If Ptr <> 0 Then Set FNodeFirstSibling = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodeLastSibling(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long, hItem As Long, hItemBuffer As Long
    hItemBuffer = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal Handle)
    Do While hItemBuffer <> 0
        hItem = hItemBuffer
        hItemBuffer = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    Loop
    If hItem = 0 Then hItem = Handle
    Ptr = GetItemPtr(hItem)
    If Ptr <> 0 Then Set FNodeLastSibling = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodeFirstVisibleSibling(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ByVal Handle))
    If Ptr <> 0 Then Set FNodeFirstVisibleSibling = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodeLastVisibleSibling(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_LASTVISIBLE, ByVal Handle))
    If Ptr <> 0 Then Set FNodeLastVisibleSibling = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodeNextVisibleSibling(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, ByVal Handle))
    If Ptr <> 0 Then Set FNodeNextVisibleSibling = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodePreviousVisibleSibling(ByVal Handle As Long) As TvwNode
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PREVIOUSVISIBLE, ByVal Handle))
    If Ptr <> 0 Then Set FNodePreviousVisibleSibling = PtrToObj(Ptr)
End If
End Property

Friend Property Get FNodeFullPath(ByVal Handle As Long) As String
If TreeViewHandle <> 0 Then
    Dim Temp As String, hItem As Long
    hItem = Handle
    Do While hItem <> 0
        Temp = PropPathSeparator & Me.FNodeText(hItem) & Temp
        hItem = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem)
    Loop
    If Not Temp = vbNullString Then If Left$(Temp, 1) = PropPathSeparator Then Temp = Mid(Temp, 2, Len(Temp) - 1)
    FNodeFullPath = Temp
End If
End Property

Friend Property Get FNodeVisible(ByVal Handle As Long) As Boolean
If TreeViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = Handle
    FNodeVisible = CBool(SendMessage(TreeViewHandle, TVM_GETITEMRECT, 0, ByVal VarPtr(RC)) <> 0)
End If
End Property

Friend Sub FNodeMove(ByVal Handle As Long, Optional ByVal RelativeNode As TvwNode, Optional ByVal Relationship As TvwNodeRelationshipConstants)
If TreeViewHandle <> 0 And Handle <> 0 Then
    Dim hParent As Long, hInsertAfter As Long, hRelative As Long
    If RelativeNode Is Nothing Then
        hRelative = TVI_ROOT
    Else
        hRelative = RelativeNode.Handle
    End If
    Select Case Relationship
        Case TvwNodeRelationshipFirst
            If Not hRelative = TVI_ROOT Then hRelative = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
            hParent = hRelative
            hInsertAfter = TVI_FIRST
        Case TvwNodeRelationshipLast
            If Not hRelative = TVI_ROOT Then hRelative = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
            hParent = hRelative
            hInsertAfter = TVI_LAST
        Case TvwNodeRelationshipNext
            If hRelative = TVI_ROOT Then
                hParent = hRelative
            Else
                hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
                If hParent = 0 Then hParent = TVI_ROOT
            End If
            hInsertAfter = hRelative
        Case TvwNodeRelationshipPrevious
            Dim hPrevious As Long
            hPrevious = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PREVIOUS, ByVal hRelative)
            If hPrevious = 0 Then
                hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
                hInsertAfter = TVI_FIRST
            Else
                hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hRelative)
                hInsertAfter = hPrevious
            End If
            If hParent = 0 Then hParent = TVI_ROOT
        Case TvwNodeRelationshipChild
            hParent = hRelative
        Case Else
            Err.Raise 380
    End Select
    If hParent <> TVI_ROOT Then
        Dim hParentTest As Long
        hParentTest = hParent
        Do While hParentTest <> 0
            If hParentTest = Handle Then Err.Raise Number:=35614, Description:="This would introduce a cycle"
            hParentTest = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hParentTest)
        Loop
    End If
    Dim Node As TvwNode
    Set Node = Me.SelectedItem
    MoveNodes Handle, hParent, hInsertAfter
    If Not Node Is Nothing Then Set Me.SelectedItem = Node
    SendMessage TreeViewHandle, TVM_DELETEITEM, 0, ByVal Handle
End If
End Sub

Friend Sub FNodeEnsureVisible(ByVal Handle As Long)
If TreeViewHandle <> 0 Then SendMessage TreeViewHandle, TVM_ENSUREVISIBLE, 0, ByVal Handle
End Sub

Friend Function FNodeCreateDragImage(ByVal Handle As Long, ByVal ImageIndex As Long) As Long
If TreeViewHandle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_IMAGE
    .hItem = Handle
    .iImage = ImageIndex - 1
    SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
    FNodeCreateDragImage = SendMessage(TreeViewHandle, TVM_CREATEDRAGIMAGE, 0, ByVal Handle)
    .iImage = I_IMAGECALLBACK
    SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI)
    End With
End If
End Function

Public Property Get SelectedNodes() As TvwSelectedNodes
Attribute SelectedNodes.VB_Description = "Returns a reference to a collection of the selected node objects."
If PropSelectedNodes Is Nothing Then
    If PropMultiSelect <> TvwMultiSelectNone Then
        Set PropSelectedNodes = New TvwSelectedNodes
        PropSelectedNodes.FInit Me
    Else
        Err.Raise Number:=91, Description:="This functionality is disabled when MultiSelect is 0 - None."
    End If
End If
Set SelectedNodes = PropSelectedNodes
End Property

Friend Function FSelectedNodesCount() As Long
FSelectedNodesCount = TreeViewSelectedCount
End Function

Friend Function FSelectedNodesItem(ByVal Index As Long) As TvwNode
' Reverse index to return the most recent selected nodes first.
' This also means that the caret item is always at index 1.
Set FSelectedNodesItem = PtrToObj(GetItemPtr(TreeViewSelectedItems(TreeViewSelectedCount - Index + 1)))
End Function

Friend Function FSelectedNodesIndex(ByVal Handle As Long) As Long
Dim i As Long
For i = 1 To TreeViewSelectedCount
    If TreeViewSelectedItems(i) = Handle Then
        ' Reverse to return correct index.
        FSelectedNodesIndex = TreeViewSelectedCount - i + 1
        Exit For
    End If
Next i
End Function

Private Sub CreateTreeView()
If TreeViewHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropRightToLeft = True Then
    If PropRightToLeftLayout = True Then
        dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    Else
        dwStyle = dwStyle Or TVS_RTLREADING
    End If
End If
Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, PropBorderStyle)
Select Case PropStyle
    Case TvwStylePlusMinusText, TvwStylePlusMinusPictureText
        dwStyle = dwStyle Or TVS_HASBUTTONS
    Case TvwStyleTreeLinesText, TvwStyleTreeLinesPictureText
        dwStyle = dwStyle Or TVS_HASLINES
    Case TvwStyleTreeLinesPlusMinusText, TvwStyleTreeLinesPlusMinusPictureText
        dwStyle = dwStyle Or TVS_HASLINES Or TVS_HASBUTTONS
End Select
If PropLineStyle = TvwLineStyleRootLines Then dwStyle = dwStyle Or TVS_LINESATROOT
If PropLabelEdit <> TvwLabelEditDisabled Then dwStyle = dwStyle Or TVS_EDITLABELS
If PropShowTips = True Then dwStyle = dwStyle Or TVS_INFOTIP
If PropHideSelection = False Then dwStyle = dwStyle Or TVS_SHOWSELALWAYS
If PropFullRowSelect = True Then dwStyle = dwStyle Or TVS_FULLROWSELECT
If PropHotTracking = True Then dwStyle = dwStyle Or TVS_TRACKSELECT
If PropScroll = False Then dwStyle = dwStyle Or TVS_NOSCROLL
If PropSingleSel = True Then dwStyle = dwStyle Or TVS_SINGLEEXPAND
If TreeViewDesignMode = False Then
    ' The WM_NOTIFYFORMAT notification must be handled, which will be sent on control creation.
    ' Thus it is necessary to subclass the parent before the control is created.
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
Else
    dwStyle = dwStyle Or TVS_NOTOOLTIPS Or TVS_DISABLEDRAGDROP
End If
TreeViewHandle = CreateWindowEx(dwExStyle, StrPtr("SysTreeView32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If TreeViewHandle <> 0 Then
    TreeViewToolTipHandle = SendMessage(TreeViewHandle, TVM_GETTOOLTIPS, 0, ByVal 0&)
    If TreeViewToolTipHandle <> 0 Then Call ComCtlsInitToolTip(TreeViewToolTipHandle)
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.BackColor = PropBackColor
Me.ForeColor = PropForeColor
If PropRedraw = False Then Me.Redraw = False
Me.LineColor = PropLineColor
' According to MSDN:
' The TVS_CHECKBOXES style must be set with SetWindowLong after the tree view control is created.
Me.Checkboxes = PropCheckboxes
Me.InsertMarkColor = PropInsertMarkColor
Me.DoubleBuffer = PropDoubleBuffer
If TreeViewHandle <> 0 Then
    If ComCtlsSupportLevel() = 0 Then
        ' According to MSDN:
        ' - If you change the font by returning CDRF_NEWFONT, the tree view control might display clipped text.
        '   This behavior is necessary for backward compatibility with earlier versions of the common controls.
        '   If you want to change the font of a tree view control, you will get better results if you send a CCM_SETVERSION message
        '   with the wParam value set to 5 before adding any items to the control.
        SendMessage TreeViewHandle, CCM_SETVERSION, 5, ByVal 0&
    End If
    SendMessage TreeViewHandle, TVM_SETINDENT, PropIndentation, ByVal 0&
End If
If TreeViewDesignMode = False Then
    If TreeViewHandle <> 0 Then
        Call ComCtlsSetSubclass(TreeViewHandle, Me, 1)
        Call ComCtlsCreateIMC(TreeViewHandle, TreeViewIMCHandle)
    End If
End If
End Sub

Private Sub DestroyTreeView()
If TreeViewHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(TreeViewHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
Call ComCtlsDestroyIMC(TreeViewHandle, TreeViewIMCHandle)
Dim hImageList As Long
hImageList = SendMessage(TreeViewHandle, TVM_GETIMAGELIST, TVSIL_STATE, ByVal 0&)
If hImageList <> 0 Then ImageList_Destroy hImageList
ShowWindow TreeViewHandle, SW_HIDE
SetParent TreeViewHandle, 0
DestroyWindow TreeViewHandle
TreeViewHandle = 0
TreeViewToolTipHandle = 0
If TreeViewFontHandle <> 0 Then
    DeleteObject TreeViewFontHandle
    TreeViewFontHandle = 0
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
If PropRedraw = True Or TreeViewDesignMode = True Then RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Function HitTest(ByVal X As Single, ByVal Y As Single) As TvwNode
Attribute HitTest.VB_Description = "Returns a reference to the node object located at the coordinates of X and Y."
If TreeViewHandle <> 0 Then
    Dim TVHTI As TVHITTESTINFO
    With TVHTI
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    If SendMessage(TreeViewHandle, TVM_HITTEST, 0, ByVal VarPtr(TVHTI)) <> 0 Then
        If (.Flags And TVHT_ONITEM) <> 0 Then
            Dim Ptr As Long
            Ptr = GetItemPtr(.hItem)
            If Ptr <> 0 Then Set HitTest = PtrToObj(Ptr)
        End If
    End If
    End With
End If
End Function

Public Function HitTestInsertMark(ByVal X As Single, ByVal Y As Single, Optional ByRef After As Boolean) As TvwNode
Attribute HitTestInsertMark.VB_Description = "Returns a reference to the node object located at the coordinates of X and Y and retrieves a value that determines where the insertion point should appear."
If TreeViewHandle <> 0 Then
    Dim TVHTI As TVHITTESTINFO
    With TVHTI
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    If SendMessage(TreeViewHandle, TVM_HITTEST, 0, ByVal VarPtr(TVHTI)) <> 0 Then
        If (.Flags And TVHT_ONITEM) <> 0 Then
            Dim Ptr As Long
            Ptr = GetItemPtr(.hItem)
            If Ptr <> 0 Then Set HitTestInsertMark = PtrToObj(Ptr)
            Dim RC As RECT
            RC.Left = .hItem
            SendMessage TreeViewHandle, TVM_GETITEMRECT, 0, ByVal VarPtr(RC)
            After = CBool(.PT.Y > (RC.Top + (RC.Bottom - RC.Top) \ 2))
        End If
    End If
    End With
End If
End Function

Public Function GetVisibleCount() As Long
Attribute GetVisibleCount.VB_Description = "Returns the number of fully visible nodes."
If TreeViewHandle <> 0 Then GetVisibleCount = SendMessage(TreeViewHandle, TVM_GETVISIBLECOUNT, 0, ByVal 0&)
End Function

Public Property Get TopItem() As TvwNode
Attribute TopItem.VB_Description = "Returns/sets a reference to the topmost visible node."
Attribute TopItem.VB_MemberFlags = "400"
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ByVal 0&))
    If Ptr <> 0 Then Set TopItem = PtrToObj(Ptr)
End If
End Property

Public Property Let TopItem(ByVal Value As TvwNode)
Set Me.TopItem = Value
End Property

Public Property Set TopItem(ByVal Value As TvwNode)
If TreeViewHandle <> 0 Then
    If Not Value Is Nothing Then
        SendMessage TreeViewHandle, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, ByVal Value.Handle
    Else
        Err.Raise 380
    End If
End If
End Property

Public Property Get SelectedItem() As TvwNode
Attribute SelectedItem.VB_Description = "Returns/sets a reference to the currently selected node."
Attribute SelectedItem.VB_MemberFlags = "400"
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&))
    If Ptr <> 0 Then Set SelectedItem = PtrToObj(Ptr)
End If
End Property

Public Property Let SelectedItem(ByVal Value As TvwNode)
Set Me.SelectedItem = Value
End Property

Public Property Set SelectedItem(ByVal Value As TvwNode)
If TreeViewHandle <> 0 Then
    If Not Value Is Nothing Then
        SendMessage TreeViewHandle, TVM_SELECTITEM, TVGN_CARET, ByVal Value.Handle
    Else
        SendMessage TreeViewHandle, TVM_SELECTITEM, TVGN_CARET, ByVal 0&
    End If
End If
End Property

Public Sub StartLabelEdit()
Attribute StartLabelEdit.VB_Description = "Begins a label editing operation on a node. This method will fail if the label edit property is set to disabled."
If TreeViewHandle <> 0 Then
    Dim hItem As Long
    hItem = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    If hItem <> 0 Then
        TreeViewStartLabelEdit = True
        SendMessage TreeViewHandle, TVM_EDITLABEL, 0, ByVal hItem
        TreeViewStartLabelEdit = False
    End If
End If
End Sub

Public Sub EndLabelEdit(Optional ByVal Discard As Boolean)
Attribute EndLabelEdit.VB_Description = "Ends the label editing operation on a node."
If TreeViewHandle <> 0 Then SendMessage TreeViewHandle, TVM_ENDEDITLABELNOW, IIf(Discard = True, 1, 0), ByVal 0&
End Sub

Public Property Get LineHeight() As Single
Attribute LineHeight.VB_Description = "Returns/sets the line height."
Attribute LineHeight.VB_MemberFlags = "400"
If TreeViewHandle <> 0 Then LineHeight = UserControl.ScaleY(SendMessage(TreeViewHandle, TVM_GETITEMHEIGHT, 0, ByVal 0&), vbPixels, vbContainerSize)
End Property

Public Property Let LineHeight(ByVal Value As Single)
If Value < 0 And Not Value = -1 Then Err.Raise 380
Dim LngValue As Long
If Value = -1 Then
    LngValue = -1
Else
    LngValue = CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
End If
If TreeViewHandle <> 0 Then SendMessage TreeViewHandle, TVM_SETITEMHEIGHT, LngValue, ByVal 0&
End Property

Public Property Get DropHighlight() As TvwNode
Attribute DropHighlight.VB_Description = "Returns/sets a reference to a node and highlights it with the system highlight color."
Attribute DropHighlight.VB_MemberFlags = "400"
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_DROPHILITE, ByVal 0&))
    If Ptr <> 0 Then Set DropHighlight = PtrToObj(Ptr)
End If
End Property

Public Property Let DropHighlight(ByVal Value As TvwNode)
Set Me.DropHighlight = Value
End Property

Public Property Set DropHighlight(ByVal Value As TvwNode)
If TreeViewHandle <> 0 Then
    If Not Value Is Nothing Then
        SendMessage TreeViewHandle, TVM_SELECTITEM, TVGN_DROPHILITE, ByVal Value.Handle
    Else
        SendMessage TreeViewHandle, TVM_SELECTITEM, TVGN_DROPHILITE, ByVal 0&
    End If
End If
End Property

Public Property Get InsertMark(Optional ByRef After As Boolean) As TvwNode
Attribute InsertMark.VB_Description = "Returns/sets a reference to a node where an insertion mark is positioned."
Attribute InsertMark.VB_MemberFlags = "400"
Dim Ptr As Long
Ptr = GetItemPtr(TreeViewInsertMarkItem)
If Ptr <> 0 Then Set InsertMark = PtrToObj(Ptr)
After = TreeViewInsertMarkAfter
End Property

Public Property Let InsertMark(Optional ByRef After As Boolean, ByVal Value As TvwNode)
Set Me.InsertMark(After) = Value
End Property

Public Property Set InsertMark(Optional ByRef After As Boolean, ByVal Value As TvwNode)
If TreeViewHandle <> 0 Then
    If Value Is Nothing Then
        TreeViewInsertMarkItem = 0
        TreeViewInsertMarkAfter = False
    Else
        If TreeViewInsertMarkItem = Value.Handle And TreeViewInsertMarkAfter = After Then Exit Property
        TreeViewInsertMarkItem = Value.Handle
        TreeViewInsertMarkAfter = After
    End If
    SendMessage TreeViewHandle, TVM_SETINSERTMARK, IIf(TreeViewInsertMarkAfter = True, 1, 0), ByVal TreeViewInsertMarkItem
End If
End Property

Public Property Get AnchorItem() As TvwNode
Attribute AnchorItem.VB_Description = "Returns/sets a reference to the anchor item. That is the item from which a multiple selection starts."
Attribute AnchorItem.VB_MemberFlags = "400"
Dim Ptr As Long
Ptr = GetItemPtr(TreeViewAnchorItem)
If Ptr <> 0 Then Set AnchorItem = PtrToObj(Ptr)
End Property

Public Property Let AnchorItem(ByVal Value As TvwNode)
Set Me.AnchorItem = Value
End Property

Public Property Set AnchorItem(ByVal Value As TvwNode)
If Not Value Is Nothing Then
    TreeViewAnchorItem = Value.Handle
Else
    TreeViewAnchorItem = 0
End If
End Property

Public Property Get OLEDraggedItem() As TvwNode
Attribute OLEDraggedItem.VB_Description = "Returns a reference to the currently dragged node during an OLE drag/drop operation."
Attribute OLEDraggedItem.VB_MemberFlags = "400"
If TreeViewDragItem <> 0 Then
    Dim Ptr As Long
    Ptr = GetItemPtr(TreeViewDragItem)
    If Ptr <> 0 Then Set OLEDraggedItem = PtrToObj(Ptr)
End If
End Property

Public Sub ResetForeColors()
Attribute ResetForeColors.VB_Description = "Resets the foreground color of particular nodes that have been modified."
If TreeViewHandle <> 0 Then
    Dim Node As TvwNode
    SendMessage TreeViewHandle, WM_SETREDRAW, 0, ByVal 0&
    For Each Node In Me.Nodes
        Node.ForeColor = -1
    Next Node
    If PropRedraw = True Then SendMessage TreeViewHandle, WM_SETREDRAW, 1, ByVal 0&
End If
End Sub

Private Sub SetVisualStylesToolTip()
If TreeViewHandle <> 0 Then
    If TreeViewToolTipHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles TreeViewToolTipHandle
        Else
            RemoveVisualStyles TreeViewToolTipHandle
        End If
    End If
End If
End Sub

Private Function GetItemPtr(ByVal Handle As Long) As Long
If Handle <> 0 Then
    Dim TVI As TVITEM
    With TVI
    .Mask = TVIF_HANDLE Or TVIF_PARAM
    .hItem = Handle
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
    GetItemPtr = .lParam
    End With
End If
End Function

Private Function MoveNodes(ByVal Handle As Long, ByVal ParentHandle As Long, ByVal InsertAfter As Long) As Long
If TreeViewHandle <> 0 And Handle <> 0 And ParentHandle <> 0 Then
    Dim TVIS As TVINSERTSTRUCT
    With TVIS
    .hParent = ParentHandle
    .hInsertAfter = InsertAfter
    With .Item.TVI
    .Mask = TVIF_HANDLE Or TVIF_TEXT Or TVIF_STATE Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM Or TVIF_INTEGRAL
    .StateMask = TVIS_FOCUSED Or TVIS_SELECTED Or TVIS_CUT Or TVIS_DROPHILITED Or TVIS_BOLD Or TVIS_EXPANDED Or TVIS_EXPANDEDONCE Or TVIS_OVERLAYMASK Or TVIS_STATEIMAGEMASK
    .hItem = Handle
    Dim Buffer As String
    Buffer = String(260, vbNullChar)
    .pszText = StrPtr(Buffer)
    .cchTextMax = 260
    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVIS.Item)
    If TreeViewSampleMode = False Then
        .iImage = I_IMAGECALLBACK
        .iSelectedImage = I_IMAGECALLBACK
    End If
    End With
    If ComCtlsSupportLevel() >= 2 Then
        Dim TVI_V61 As TVITEMEX_V61
        TVI_V61.TVI.Mask = TVIF_HANDLE Or TVIF_STATEEX
        TVI_V61.TVI.hItem = Handle
        SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI_V61)
        MoveNodes = SendMessage(TreeViewHandle, TVM_INSERTITEM, 0, ByVal VarPtr(TVIS))
        TVI_V61.TVI.hItem = MoveNodes
        SendMessage TreeViewHandle, TVM_SETITEM, 0, ByVal VarPtr(TVI_V61)
    Else
        MoveNodes = SendMessage(TreeViewHandle, TVM_INSERTITEM, 0, ByVal VarPtr(TVIS))
    End If
    If MoveNodes <> 0 Then
        Dim Node As TvwNode
        Set Node = PtrToObj(.Item.TVI.lParam)
        Node.Handle = MoveNodes
        Dim hChild As Long
        hChild = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CHILD, ByVal Handle)
        Do While hChild <> 0
            MoveNodes hChild, MoveNodes, TVI_LAST
            hChild = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hChild)
        Loop
    End If
    End With
End If
End Function

Private Sub SortNodes(ByVal Handle As Long, ByVal SortType As TvwSortTypeConstants)
If TreeViewHandle <> 0 And Handle <> 0 Then
    Dim TVSCB As TVSORTCB
    With TVSCB
    .hParent = Handle
    Select Case SortType
        Case TvwSortTypeBinary
            .lpfnCompare = ProcPtr(AddressOf ComCtlsTvwSortingFunctionBinary)
        Case TvwSortTypeText
            .lpfnCompare = ProcPtr(AddressOf ComCtlsTvwSortingFunctionText)
    End Select
    If .lpfnCompare <> 0 Then
        Dim This As ISubclass
        Set This = Me
        .lParam = ObjPtr(This)
        SendMessage TreeViewHandle, TVM_SORTCHILDRENCB, 0, ByVal VarPtr(TVSCB)
    End If
    End With
End If
End Sub

Private Sub NodesRemoveRecursion(ByVal ChildNodes As Collection, ByVal hChild As Long)
If TreeViewHandle <> 0 Then
    Dim Ptr As Long
    hChild = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hChild)
    Do While hChild <> 0
        Ptr = GetItemPtr(hChild)
        If Ptr <> 0 Then ChildNodes.Add Ptr
        Call NodesRemoveRecursion(ChildNodes, hChild)
        hChild = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hChild)
    Loop
End If
End Sub

Private Function GetSelectRange(ByVal hItem1 As Long, ByVal hItem2 As Long) As Collection
Set GetSelectRange = New Collection
If TreeViewHandle <> 0 Then
    If hItem1 <> 0 And hItem2 <> 0 Then
        GetSelectRange.Add hItem1
        Dim hParent As Long
        hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem2)
        Dim hItem As Long, RC As RECT
        hItem = hItem1
        Do Until hItem = hItem2
            hItem = GetSelectRangeStep(hItem)
            If hItem = 0 Then Exit Do
            If PropMultiSelect = TvwMultiSelectAll Then
                GetSelectRange.Add hItem
            ElseIf PropMultiSelect = TvwMultiSelectVisibleOnly Then
                RC.Left = hItem
                If SendMessage(TreeViewHandle, TVM_GETITEMRECT, 0, ByVal VarPtr(RC)) <> 0 Then GetSelectRange.Add hItem
            ElseIf PropMultiSelect = TvwMultiSelectRestrictSiblings Then
                If SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem) = hParent Then GetSelectRange.Add hItem
            End If
        Loop
        If hItem = 0 Then
            Set GetSelectRange = New Collection
            GetSelectRange.Add hItem2
            hItem = hItem2
            Do Until hItem = hItem1
                hItem = GetSelectRangeStep(hItem)
                If hItem = 0 Then Exit Do
                If PropMultiSelect = TvwMultiSelectAll Then
                    GetSelectRange.Add hItem, , 1
                ElseIf PropMultiSelect = TvwMultiSelectVisibleOnly Then
                    RC.Left = hItem
                    If SendMessage(TreeViewHandle, TVM_GETITEMRECT, 0, ByVal VarPtr(RC)) <> 0 Then GetSelectRange.Add hItem, , 1
                ElseIf PropMultiSelect = TvwMultiSelectRestrictSiblings Then
                    If SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem) = hParent Then GetSelectRange.Add hItem, , 1
                End If
            Loop
        End If
    ElseIf hItem2 <> 0 Then
        GetSelectRange.Add hItem2
    End If
End If
End Function

Private Function GetSelectRangeStep(ByVal hItem As Long) As Long
If TreeViewHandle <> 0 And hItem <> 0 Then
    Dim hStep As Long
    hStep = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hItem)
    If hStep <> 0 Then
        GetSelectRangeStep = hStep
        Exit Function
    End If
    hStep = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    If hStep <> 0 Then
        GetSelectRangeStep = hStep
        Exit Function
    End If
    hStep = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem)
    If hStep = 0 Then Exit Function
    hStep = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hStep)
    If hStep <> 0 Then
        GetSelectRangeStep = hStep
    Else
        Dim hParent As Long
        hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem)
        While hParent <> 0
            hStep = hParent
            hParent = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hParent)
        Wend
        hStep = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hStep)
        If hStep <> 0 Then GetSelectRangeStep = hStep
    End If
End If
End Function

Private Sub SetItemSelected(ByVal Handle As Long, ByVal State As Boolean, Optional ByVal NoRedraw As Boolean)
Dim i As Long
For i = 1 To TreeViewSelectedCount
    If TreeViewSelectedItems(i) = Handle Then Exit For
Next i
If State = True Then
    If i > TreeViewSelectedCount Then
        TreeViewSelectedCount = TreeViewSelectedCount + 1
        ReDim Preserve TreeViewSelectedItems(1 To TreeViewSelectedCount) As Long
        TreeViewSelectedItems(TreeViewSelectedCount) = Handle
    Else
        NoRedraw = True
    End If
Else
    If i <= TreeViewSelectedCount And TreeViewSelectedCount > 0 Then
        Dim j As Long
        For j = i To TreeViewSelectedCount - 1
            TreeViewSelectedItems(j) = TreeViewSelectedItems(j + 1)
        Next j
        TreeViewSelectedCount = TreeViewSelectedCount - 1
        If TreeViewSelectedCount > 0 Then
            ReDim Preserve TreeViewSelectedItems(1 To TreeViewSelectedCount) As Long
        Else
            Erase TreeViewSelectedItems()
        End If
    Else
        NoRedraw = True
    End If
End If
If NoRedraw = False Then
    If TreeViewHandle <> 0 Then
        Dim RC As RECT
        RC.Left = Handle
        If SendMessage(TreeViewHandle, TVM_GETITEMRECT, 0, ByVal VarPtr(RC)) <> 0 Then
            InvalidateRect TreeViewHandle, RC, 1
            UpdateWindow TreeViewHandle
        End If
    End If
End If
End Sub

Private Function IsItemSelected(ByVal Handle As Long) As Boolean
Dim i As Long
For i = 1 To TreeViewSelectedCount
    If TreeViewSelectedItems(i) = Handle Then Exit For
Next i
If i <= TreeViewSelectedCount And TreeViewSelectedCount > 0 Then IsItemSelected = True
End Function

Private Function ClearSelectedItems(Optional ByVal Handle As Variant) As Long
Dim RC As RECT, NeedUpdate As Boolean
If TreeViewSelectedCount > 0 Then
    If TreeViewHandle <> 0 Then
        Dim i As Long
        For i = 1 To TreeViewSelectedCount
            RC.Left = TreeViewSelectedItems(i)
            If SendMessage(TreeViewHandle, TVM_GETITEMRECT, 0, ByVal VarPtr(RC)) <> 0 Then
                InvalidateRect TreeViewHandle, RC, 1
                NeedUpdate = True
            End If
        Next i
    End If
    TreeViewSelectedCount = 0
    Erase TreeViewSelectedItems()
End If
If TreeViewHandle <> 0 Then
    If PropMultiSelect <> TvwMultiSelectNone Then
        Dim hItem As Long
        If IsMissing(Handle) Then
            hItem = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
        Else
            hItem = Handle
        End If
        If hItem <> 0 Then
            ClearSelectedItems = hItem
            If (SendMessage(TreeViewHandle, TVM_GETITEMSTATE, hItem, ByVal TVIS_SELECTED) And TVIS_SELECTED) <> 0 Then
                TreeViewSelectedCount = 1
                ReDim Preserve TreeViewSelectedItems(1 To TreeViewSelectedCount) As Long
                TreeViewSelectedItems(TreeViewSelectedCount) = hItem
                RC.Left = TreeViewSelectedItems(1)
                If SendMessage(TreeViewHandle, TVM_GETITEMRECT, 0, ByVal VarPtr(RC)) <> 0 Then
                    InvalidateRect TreeViewHandle, RC, 1
                    NeedUpdate = True
                End If
            End If
        End If
    End If
    If NeedUpdate = True Then UpdateWindow TreeViewHandle
End If
End Function

Private Function NodesSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String, Node1 As TvwNode, Node2 As TvwNode, ParentNode As TvwNode
Set Node1 = PtrToObj(lParam1)
Set Node2 = PtrToObj(lParam2)
Text1 = Node1.Text
Text2 = Node2.Text
NodesSortingFunctionBinary = lstrcmp(StrPtr(Text1), StrPtr(Text2))
Set ParentNode = Node1.Parent
If ParentNode Is Nothing Then
    If PropSortOrder = TvwSortOrderDescending Then NodesSortingFunctionBinary = -NodesSortingFunctionBinary
Else
    If ParentNode.SortOrder = TvwSortOrderDescending Then NodesSortingFunctionBinary = -NodesSortingFunctionBinary
End If
End Function

Private Function NodesSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String, Node1 As TvwNode, Node2 As TvwNode, ParentNode As TvwNode
Set Node1 = PtrToObj(lParam1)
Set Node2 = PtrToObj(lParam2)
Text1 = Node1.Text
Text2 = Node2.Text
NodesSortingFunctionText = lstrcmpi(StrPtr(Text1), StrPtr(Text2))
Set ParentNode = Node1.Parent
If ParentNode Is Nothing Then
    If PropSortOrder = TvwSortOrderDescending Then NodesSortingFunctionText = -NodesSortingFunctionText
Else
    If ParentNode.SortOrder = TvwSortOrderDescending Then NodesSortingFunctionText = -NodesSortingFunctionText
End If
End Function

Private Function IndexToStateImageMask(ByVal ImgIndex As Long) As Long
IndexToStateImageMask = ImgIndex * (2 ^ 12)
End Function

Private Function StateImageMaskToIndex(ByVal ImgState As Long) As Long
StateImageMaskToIndex = ImgState / (2 ^ 12)
End Function

Private Function PropImageListControl() As Object
If TreeViewImageListObjectPointer <> 0 Then Set PropImageListControl = PtrToObj(TreeViewImageListObjectPointer)
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcLabelEdit(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Case 10
        ISubclass_Message = NodesSortingFunctionBinary(wParam, lParam)
    Case 11
        ISubclass_Message = NodesSortingFunctionText(wParam, lParam)
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
        PostMessage hWnd, UM_BUTTONDOWN, MakeDWord(vbLeftButton, GetShiftStateFromParam(wParam)), ByVal lParam
    Case WM_MBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
        ' There is no modal message loop (DragDetect) on WM_MBUTTONDOWN.
    Case WM_RBUTTONDOWN
        If GetFocus() <> hWnd Then UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
        PostMessage hWnd, UM_BUTTONDOWN, MakeDWord(vbRightButton, GetShiftStateFromParam(wParam)), ByVal lParam
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
                If KeyCode = vbKeySpace And PropCheckboxes = True Then
                    Dim TVI As TVITEM
                    With TVI
                    .Mask = TVIF_HANDLE Or TVIF_STATE Or TVIF_PARAM
                    .hItem = SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
                    .StateMask = TVIS_STATEIMAGEMASK
                    SendMessage TreeViewHandle, TVM_GETITEM, 0, ByVal VarPtr(TVI)
                    If StateImageMaskToIndex(.State And TVIS_STATEIMAGEMASK) = 0 Then Exit Function
                    If .lParam <> 0 Then
                        Dim Cancel As Boolean
                        RaiseEvent NodeBeforeCheck(PtrToObj(.lParam), Cancel)
                        If Cancel = True Then Exit Function
                    End If
                    PostMessage TreeViewHandle, UM_CHECKSTATECHANGED, 0, ByVal .lParam
                    End With
                End If
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            TreeViewCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If TreeViewCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(TreeViewCharCodeCache And &HFFFF&)
            TreeViewCharCodeCache = 0
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
    Case WM_INPUTLANGCHANGE
        Call ComCtlsSetIMEMode(hWnd, TreeViewIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call ComCtlsSetIMEMode(hWnd, TreeViewIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case UM_CHECKSTATECHANGED
        If lParam <> 0 Then RaiseEvent NodeCheck(PtrToObj(lParam))
        Exit Function
    Case UM_BUTTONDOWN
        ' The control enters a modal message loop (DragDetect) on WM_LBUTTONDOWN and WM_RBUTTONDOWN.
        ' This workaround is necessary to raise 'MouseDown' before the button was released or the mouse was moved.
        RaiseEvent MouseDown(LoWord(wParam), HiWord(wParam), UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips), UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips))
        TreeViewButtonDown = LoWord(wParam)
        TreeViewIsClick = True
        Exit Function
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_SETFOCUS, WM_KILLFOCUS
        TreeViewFocused = CBool(wMsg = WM_SETFOCUS)
        If PropMultiSelect <> TvwMultiSelectNone Then Me.Refresh
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                ' In case DragDetect returns 0 then the control will set focus the focus automatically.
                ' Otherwise not. So check and change focus, if needed.
                If GetFocus() <> hWnd Then SetFocusAPI hWnd
                ' See UM_BUTTONDOWN
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                TreeViewButtonDown = 0
                TreeViewIsClick = True
            Case WM_RBUTTONDOWN
                ' In case DragDetect returns 0 then the control will set focus the focus automatically.
                ' Otherwise not. So check and change focus, if needed.
                If GetFocus() <> hWnd Then SetFocusAPI hWnd
                ' See UM_BUTTONDOWN
            Case WM_MOUSEMOVE
                If TreeViewMouseOver = False And PropMouseTrack = True Then
                    TreeViewMouseOver = True
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
                TreeViewButtonDown = 0
                If TreeViewIsClick = True Then
                    TreeViewIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If TreeViewMouseOver = True Then
            TreeViewMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcLabelEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_KEYDOWN, WM_KEYUP
        TreeViewCharCodeCache = ComCtlsPeekCharCode(hWnd)
    Case WM_CHAR
        If TreeViewCharCodeCache <> 0 Then
            wParam = TreeViewCharCodeCache
            TreeViewCharCodeCache = 0
        End If
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            WindowProcLabelEdit = 1
        Else
            Dim UTF16 As String
            UTF16 = UTF32CodePoint_To_UTF16(wParam)
            If Len(UTF16) = 1 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(UTF16)), ByVal lParam
            ElseIf Len(UTF16) = 2 Then
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Left$(UTF16, 1))), ByVal lParam
                SendMessage hWnd, WM_CHAR, CIntToUInt(AscW(Right$(UTF16, 1))), ByVal lParam
            End If
            WindowProcLabelEdit = 0
        End If
        Exit Function
    Case WM_INPUTLANGCHANGE
        Call ComCtlsSetIMEMode(hWnd, TreeViewIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call ComCtlsSetIMEMode(hWnd, TreeViewIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
End Select
WindowProcLabelEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = TreeViewHandle Then
            Dim Ptr As Long, Node As TvwNode, hEnum As Variant
            Dim Length As Long, Cancel As Boolean
            Dim NMTV As NMTREEVIEW, NMTVDI As NMTVDISPINFO, TVHTI As TVHITTESTINFO
            Select Case NM.Code
                Case TVN_BEGINLABELEDIT, TVN_ENDLABELEDIT
                    Static LabelEditHandle As Long
                    Select Case NM.Code
                        Case TVN_BEGINLABELEDIT
                            If PropLabelEdit = TvwLabelEditManual And TreeViewStartLabelEdit = False Then
                                WindowProcUserControl = 1
                            Else
                                If PropMultiSelect <> TvwMultiSelectNone Then
                                    ' Suppress label edits in a multi select tree view under certain conditions.
                                    If TreeViewStartLabelEdit = True Then
                                        ' Never suppress a label edit initiated by code.
                                    ElseIf TreeViewClickSelectedCount <> 1 Or (TreeViewClickShift And (vbShiftMask Or vbCtrlMask)) <> 0 Then
                                        Cancel = True
                                    End If
                                End If
                                If Cancel = False Then
                                    RaiseEvent BeforeLabelEdit(Cancel)
                                    If Cancel = True Then
                                        WindowProcUserControl = 1
                                    Else
                                        WindowProcUserControl = 0
                                        LabelEditHandle = Me.hWndLabelEdit
                                        If LabelEditHandle <> 0 Then
                                            If PropRightToLeft = True And PropRightToLeftLayout = False Then Call ComCtlsSetRightToLeft(LabelEditHandle, WS_EX_RTLREADING)
                                            Call ComCtlsSetSubclass(LabelEditHandle, Me, 2)
                                        End If
                                        TreeViewLabelInEdit = True
                                    End If
                                Else
                                    WindowProcUserControl = 1
                                End If
                            End If
                        Case TVN_ENDLABELEDIT
                            CopyMemory NMTVDI, ByVal lParam, LenB(NMTVDI)
                            With NMTVDI.Item
                            If .pszText <> 0 Then
                                Dim NewText As String
                                Length = lstrlen(.pszText)
                                If Length > 0 Then
                                    NewText = String(Length, vbNullChar)
                                    CopyMemory ByVal StrPtr(NewText), ByVal .pszText, Length * 2
                                End If
                                RaiseEvent AfterLabelEdit(Cancel, NewText)
                                If Cancel = False Then
                                    WindowProcUserControl = 1
                                Else
                                    WindowProcUserControl = 0
                                End If
                            Else
                                WindowProcUserControl = 0
                            End If
                            End With
                            If LabelEditHandle <> 0 Then
                                Call ComCtlsRemoveSubclass(LabelEditHandle)
                                LabelEditHandle = 0
                            End If
                            TreeViewLabelInEdit = False
                    End Select
                    Exit Function
                Case TVN_BEGINDRAG, TVN_BEGINRDRAG
                    CopyMemory NMTV, ByVal lParam, LenB(NMTV)
                    With NMTV.ItemNew
                    If .lParam <> 0 Then
                        Set Node = PtrToObj(.lParam)
                        TreeViewDragItemBuffer = .hItem
                        If NM.Code = TVN_BEGINDRAG Then
                            RaiseEvent NodeDrag(Node, vbLeftButton)
                            If PropOLEDragMode = vbOLEDragAutomatic Then Me.OLEDrag
                        ElseIf NM.Code = TVN_BEGINRDRAG Then
                            RaiseEvent NodeDrag(Node, vbRightButton)
                        End If
                        TreeViewDragItemBuffer = 0
                    End If
                    End With
                Case NM_CLICK, NM_RCLICK
                    With TVHTI
                    GetCursorPos .PT
                    ScreenToClient TreeViewHandle, .PT
                    SendMessage TreeViewHandle, TVM_HITTEST, 0, ByVal VarPtr(TVHTI)
                    If .hItem <> 0 Then
                        If PropMultiSelect <> TvwMultiSelectNone Then
                            TreeViewClickSelectedCount = TreeViewSelectedCount
                            TreeViewClickShift = GetShiftStateFromMsg()
                            ' TVN_SELCHANGED will not be fired when a click is on the current focused item.
                            ' Ensure the click is on the label or icon and would normally cause a TVN_SELCHANGED.
                            If (.Flags And (TVHT_ONITEMICON Or TVHT_ONITEMLABEL)) <> 0 Then
                                If SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&) = .hItem Then
                                    If (TreeViewClickShift And (vbShiftMask Or vbCtrlMask)) = 0 Then
                                        ' Clear all highlighted items in case it is a simple click.
                                        TreeViewAnchorItem = ClearSelectedItems()
                                        Me.FNodeSelected(.hItem) = True
                                    Else
                                        If (TreeViewClickShift And vbShiftMask) = 0 And (TreeViewClickShift And vbCtrlMask) <> 0 Then
                                            ' Toggle highlighted state in case it is just a control click.
                                            Me.FNodeSelected(.hItem) = Not IsItemSelected(.hItem)
                                        ElseIf (TreeViewClickShift And vbShiftMask) <> 0 Then
                                            If TreeViewAnchorItem <> .hItem Then
                                                If (TreeViewClickShift And vbCtrlMask) = 0 Then ClearSelectedItems 0&
                                                For Each hEnum In GetSelectRange(TreeViewAnchorItem, .hItem)
                                                    Select Case hEnum
                                                        Case TreeViewAnchorItem, .hItem
                                                            Me.FNodeSelected(hEnum) = True
                                                        Case Else
                                                            Cancel = False
                                                            RaiseEvent NodeRangeSelect(PtrToObj(GetItemPtr(hEnum)), Cancel)
                                                            If Cancel = False Then Me.FNodeSelected(hEnum) = True
                                                    End Select
                                                Next hEnum
                                            Else
                                                ClearSelectedItems .hItem
                                                Me.FNodeSelected(.hItem) = True
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf (.Flags And TVHT_ONITEMRIGHT) <> 0 Then
                                If (TreeViewClickShift And (vbShiftMask Or vbCtrlMask)) = 0 Then
                                    ' Clear all highlighted items in case it is a simple click.
                                    TreeViewAnchorItem = ClearSelectedItems()
                                End If
                            End If
                        End If
                        Ptr = GetItemPtr(.hItem)
                        If Ptr <> 0 Then Set Node = PtrToObj(Ptr)
                        If (.Flags And (TVHT_ONITEMICON Or TVHT_ONITEMLABEL)) <> 0 Then
                            If NM.Code = NM_CLICK Then
                                RaiseEvent NodeClick(Node, vbLeftButton)
                            ElseIf NM.Code = NM_RCLICK Then
                                RaiseEvent NodeClick(Node, vbRightButton)
                            End If
                        End If
                        If ((.Flags And TVHT_ONITEMBUTTON) = 0 Or TreeViewButtonDown <> vbLeftButton) And TreeViewButtonDown <> 0 Then
                            RaiseEvent MouseUp(TreeViewButtonDown, GetShiftStateFromMsg(), UserControl.ScaleX(.PT.X, vbPixels, vbTwips), UserControl.ScaleY(.PT.Y, vbPixels, vbTwips))
                            TreeViewButtonDown = 0
                            TreeViewIsClick = False
                            RaiseEvent Click
                        End If
                        If (.Flags And TVHT_ONITEMSTATEICON) = TVHT_ONITEMSTATEICON And PropCheckboxes = True Then
                            RaiseEvent NodeBeforeCheck(Node, Cancel)
                            If Cancel = True Then
                                WindowProcUserControl = 1
                                Exit Function
                            Else
                                PostMessage TreeViewHandle, UM_CHECKSTATECHANGED, 0, ByVal Ptr
                            End If
                        End If
                    Else
                        If PropMultiSelect <> TvwMultiSelectNone Then
                            TreeViewClickSelectedCount = TreeViewSelectedCount
                            TreeViewClickShift = GetShiftStateFromMsg()
                            ' TVN_SELCHANGED will not be fired when a click is nowhere.
                            If (TreeViewClickShift And (vbShiftMask Or vbCtrlMask)) = 0 Then
                                ' Clear all highlighted items in case it is a simple click.
                                TreeViewAnchorItem = ClearSelectedItems()
                            End If
                        End If
                        If TreeViewButtonDown <> vbLeftButton And TreeViewButtonDown <> 0 Then
                            RaiseEvent MouseUp(TreeViewButtonDown, GetShiftStateFromMsg(), UserControl.ScaleX(.PT.X, vbPixels, vbTwips), UserControl.ScaleY(.PT.Y, vbPixels, vbTwips))
                            TreeViewButtonDown = 0
                            TreeViewIsClick = False
                            RaiseEvent Click
                        End If
                    End If
                    End With
                Case NM_DBLCLK, NM_RDBLCLK
                    With TVHTI
                    GetCursorPos .PT
                    ScreenToClient TreeViewHandle, .PT
                    SendMessage TreeViewHandle, TVM_HITTEST, 0, ByVal VarPtr(TVHTI)
                    If .hItem <> 0 Then
                        Ptr = GetItemPtr(.hItem)
                        If Ptr <> 0 Then Set Node = PtrToObj(Ptr)
                        If (.Flags And (TVHT_ONITEMICON Or TVHT_ONITEMLABEL)) <> 0 Then
                            If NM.Code = NM_DBLCLK Then
                                RaiseEvent NodeDblClick(Node, vbLeftButton)
                            ElseIf NM.Code = NM_RDBLCLK Then
                                RaiseEvent NodeDblClick(Node, vbRightButton)
                            End If
                        End If
                    End If
                    End With
                    RaiseEvent DblClick
                Case TVN_ITEMEXPANDING
                    CopyMemory NMTV, ByVal lParam, LenB(NMTV)
                    With NMTV
                    If .ItemNew.lParam <> 0 Then
                        Set Node = PtrToObj(.ItemNew.lParam)
                        If (.Action And TVE_COLLAPSE) = TVE_COLLAPSE Then
                            RaiseEvent BeforeCollapse(Node, Cancel)
                        ElseIf (.Action And TVE_EXPAND) = TVE_EXPAND Then
                            RaiseEvent BeforeExpand(Node, Cancel)
                        End If
                        If Cancel = True Then
                            WindowProcUserControl = 1
                        Else
                            WindowProcUserControl = 0
                        End If
                        Exit Function
                    End If
                    End With
                Case TVN_ITEMEXPANDED
                    CopyMemory NMTV, ByVal lParam, LenB(NMTV)
                    With NMTV
                    If .ItemNew.lParam <> 0 Then
                        Set Node = PtrToObj(.ItemNew.lParam)
                        If (.Action And TVE_COLLAPSE) = TVE_COLLAPSE Then
                            If Node.ExpandedImage > 0 Then Me.FNodeRedraw .ItemNew.hItem
                            RaiseEvent Collapse(Node)
                        ElseIf (.Action And TVE_EXPAND) = TVE_EXPAND Then
                            If Node.ExpandedImage > 0 Then Me.FNodeRedraw .ItemNew.hItem
                            RaiseEvent Expand(Node)
                        End If
                    End If
                    End With
                Case TVN_SELCHANGING
                    CopyMemory NMTV, ByVal lParam, LenB(NMTV)
                    With NMTV
                    If .ItemNew.lParam <> 0 Then
                        Set Node = PtrToObj(.ItemNew.lParam)
                        RaiseEvent NodeBeforeSelect(Node, Cancel)
                        If Cancel = True Then
                            WindowProcUserControl = 1
                        Else
                            WindowProcUserControl = 0
                        End If
                        Exit Function
                    End If
                    End With
                Case TVN_SELCHANGED
                    CopyMemory NMTV, ByVal lParam, LenB(NMTV)
                    With NMTV
                    If .ItemNew.hItem <> 0 Then
                        If PropMultiSelect <> TvwMultiSelectNone Then
                            Select Case .Action
                                Case TVC_BYMOUSE, TVC_BYKEYBOARD
                                    Dim Shift As Integer
                                    Shift = GetShiftStateFromMsg()
                                    If TreeViewAnchorItem = 0 Or (Shift And vbShiftMask) = 0 Then TreeViewAnchorItem = .ItemNew.hItem
                                    If (Shift And (vbShiftMask Or vbCtrlMask)) = 0 Then
                                        ' Clear all highlighted items in case it is a simple click.
                                        ClearSelectedItems .ItemNew.hItem
                                    Else
                                        If PropMultiSelect = TvwMultiSelectRestrictSiblings Then
                                            If SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal .ItemOld.hItem) <> SendMessage(TreeViewHandle, TVM_GETNEXTITEM, TVGN_PARENT, ByVal .ItemNew.hItem) Then
                                                TreeViewAnchorItem = ClearSelectedItems(.ItemNew.hItem)
                                                Shift = 0
                                            End If
                                        End If
                                        If (Shift And vbShiftMask) = 0 And (Shift And vbCtrlMask) <> 0 Then
                                            ' Toggle highlighted state in case it is just a control click.
                                            Me.FNodeSelected(.ItemNew.hItem) = Not IsItemSelected(.ItemNew.hItem)
                                        ElseIf (Shift And vbShiftMask) <> 0 Then
                                            If TreeViewAnchorItem <> .ItemNew.hItem Then
                                                If (Shift And vbCtrlMask) = 0 Then ClearSelectedItems 0&
                                                For Each hEnum In GetSelectRange(TreeViewAnchorItem, .ItemNew.hItem)
                                                    Select Case hEnum
                                                        Case TreeViewAnchorItem, .ItemNew.hItem
                                                            Me.FNodeSelected(hEnum) = True
                                                        Case Else
                                                            Cancel = False
                                                            RaiseEvent NodeRangeSelect(PtrToObj(GetItemPtr(hEnum)), Cancel)
                                                            If Cancel = False Then Me.FNodeSelected(hEnum) = True
                                                    End Select
                                                Next hEnum
                                            Else
                                                ClearSelectedItems .ItemNew.hItem
                                            End If
                                        Else
                                            ' Fallback in case the shift variable got cleared.
                                            ClearSelectedItems .ItemNew.hItem
                                        End If
                                    End If
                                Case Else
                                    ' It is not safe to rely on TVC_UNKNOWN only. The action member can have an unknown value.
                                    ' If the caret item was changed by code then the action value is TVC_UNKNOWN for sure.
                                    ' However, if there is no caret item and the tree view received focus by tab key then the action value is &H1000.
                                    ' Thus clear all highlighted items in case the focused item was changed by code or by an unknown action value.
                                    TreeViewAnchorItem = ClearSelectedItems(.ItemNew.hItem)
                            End Select
                        End If
                        If .ItemNew.lParam <> 0 Then
                            Set Node = PtrToObj(.ItemNew.lParam)
                            RaiseEvent NodeSelect(Node)
                        End If
                    Else
                        ' If hItem is zero then there is no caret item anymore.
                        ' Clear all selected items in case the tree view is multi select.
                        If PropMultiSelect <> TvwMultiSelectNone Then TreeViewAnchorItem = ClearSelectedItems(0&)
                    End If
                    End With
                Case TVN_DELETEITEM
                    If PropMultiSelect <> TvwMultiSelectNone Then
                        CopyMemory NMTV, ByVal lParam, LenB(NMTV)
                        With NMTV.ItemOld
                        If .hItem = TreeViewAnchorItem Then TreeViewAnchorItem = 0
                        Call SetItemSelected(.hItem, False, True)
                        End With
                    End If
                Case NM_CUSTOMDRAW
                    Dim NMTVCD As NMTVCUSTOMDRAW
                    CopyMemory NMTVCD, ByVal lParam, LenB(NMTVCD)
                    Select Case NMTVCD.NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            WindowProcUserControl = CDRF_NOTIFYITEMDRAW
                            Exit Function
                        Case CDDS_ITEMPREPAINT
                            With NMTVCD
                            If .NMCD.lItemlParam <> 0 Then
                                Set Node = PtrToObj(.NMCD.lItemlParam)
                                If (.NMCD.uItemState And CDIS_FOCUS) = 0 And (.NMCD.uItemState And CDIS_SELECTED) = 0 Then
                                    ' CDIS_DROPHILITED will never be set so check for TVIS_DROPHILITED instead.
                                    If (SendMessage(TreeViewHandle, TVM_GETITEMSTATE, .NMCD.dwItemSpec, ByVal TVIS_DROPHILITED) And TVIS_DROPHILITED) = 0 Then
                                        Dim HighLighted As Boolean
                                        If PropMultiSelect <> TvwMultiSelectNone Then
                                            If IsItemSelected(.NMCD.dwItemSpec) = True Then
                                                If TreeViewFocused = True Then
                                                    .ClrText = WinColor(vbHighlightText)
                                                    .ClrTextBk = WinColor(vbHighlight)
                                                    CopyMemory ByVal lParam, NMTVCD, LenB(NMTVCD)
                                                    HighLighted = True
                                                ElseIf PropHideSelection = False Then
                                                    .ClrText = WinColor(vbWindowText)
                                                    .ClrTextBk = WinColor(vbButtonFace)
                                                    CopyMemory ByVal lParam, NMTVCD, LenB(NMTVCD)
                                                    HighLighted = True
                                                End If
                                            End If
                                        End If
                                        If HighLighted = False Then
                                            If (.NMCD.uItemState And CDIS_DISABLED) = 0 Then
                                                If (.NMCD.uItemState And CDIS_HOT) = 0 Then
                                                    .ClrText = WinColor(Node.ForeColor)
                                                Else
                                                    .ClrText = GetSysColor(COLOR_HOTLIGHT)
                                                End If
                                                .ClrTextBk = WinColor(Node.BackColor)
                                            Else
                                                .ClrText = WinColor(vbGrayText)
                                                .ClrTextBk = WinColor(vbButtonFace)
                                            End If
                                            CopyMemory ByVal lParam, NMTVCD, LenB(NMTVCD)
                                        End If
                                    End If
                                End If
                                If Node.NoImages = False Then
                                    WindowProcUserControl = CDRF_DODEFAULT
                                Else
                                    WindowProcUserControl = TVCDRF_NOIMAGES
                                End If
                            Else
                                WindowProcUserControl = CDRF_DODEFAULT
                            End If
                            End With
                            Exit Function
                    End Select
                Case TVN_GETDISPINFO
                    CopyMemory NMTVDI, ByVal lParam, LenB(NMTVDI)
                    With NMTVDI.Item
                    If .lParam <> 0 Then
                        Set Node = PtrToObj(.lParam)
                        If (.Mask And TVIF_IMAGE) = TVIF_IMAGE Then
                            If Node.ExpandedImageIndex = 0 Then
                                .iImage = Node.ImageIndex - 1
                            Else
                                If Me.FNodeExpanded(.hItem) = True Then
                                    .iImage = Node.ExpandedImageIndex - 1
                                Else
                                    .iImage = Node.ImageIndex - 1
                                End If
                            End If
                        End If
                        If (.Mask And TVIF_SELECTEDIMAGE) = TVIF_SELECTEDIMAGE Then
                            If Node.ExpandedImageIndex = 0 Then
                                .iSelectedImage = Node.SelectedImageIndex - 1
                            Else
                                If Me.FNodeExpanded(.hItem) = True Then
                                    If Node.SelectedImageIndex = 0 Then
                                        .iSelectedImage = Node.ExpandedImageIndex - 1
                                    Else
                                        .iSelectedImage = Node.SelectedImageIndex - 1
                                    End If
                                Else
                                    .iSelectedImage = Node.SelectedImageIndex - 1
                                End If
                            End If
                            If .iSelectedImage = I_IMAGECALLBACK Then .iSelectedImage = Node.ImageIndex - 1
                        End If
                    End If
                    End With
                    CopyMemory ByVal lParam, NMTVDI, LenB(NMTVDI)
                Case TVN_GETINFOTIP
                    Dim NMTVGIT As NMTVGETINFOTIP
                    CopyMemory NMTVGIT, ByVal lParam, LenB(NMTVGIT)
                    With NMTVGIT
                    If .hItem <> 0 And .lParam <> 0 And .pszText <> 0 Then
                        Set Node = PtrToObj(.lParam)
                        Dim ToolTipText As String
                        ToolTipText = Node.ToolTipText
                        If Not ToolTipText = vbNullString Then
                            ToolTipText = Left$(ToolTipText, .cchTextMax - 1) & vbNullChar
                            CopyMemory ByVal .pszText, ByVal StrPtr(ToolTipText), LenB(ToolTipText)
                        Else
                            CopyMemory ByVal .pszText, 0&, 4
                        End If
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
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI TreeViewHandle
End Function
