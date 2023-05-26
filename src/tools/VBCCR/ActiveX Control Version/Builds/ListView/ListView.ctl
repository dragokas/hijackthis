VERSION 5.00
Begin VB.UserControl ListView 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "ListView.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ListView.ctx":0074
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
#If False Then
Private LvwViewIcon, LvwViewSmallIcon, LvwViewList, LvwViewReport, LvwViewTile
Private LvwArrangeNone, LvwArrangeAutoLeft, LvwArrangeAutoTop, LvwArrangeLeft, LvwArrangeTop
Private LvwColumnHeaderAlignmentLeft, LvwColumnHeaderAlignmentRight, LvwColumnHeaderAlignmentCenter
Private LvwColumnHeaderSortArrowNone, LvwColumnHeaderSortArrowDown, LvwColumnHeaderSortArrowUp
Private LvwColumnHeaderAutoSizeToItems, LvwColumnHeaderAutoSizeToHeader
Private LvwColumnHeaderFilterTypeText, LvwColumnHeaderFilterTypeNumber
Private LvwLabelEditAutomatic, LvwLabelEditManual, LvwLabelEditDisabled
Private LvwSortOrderAscending, LvwSortOrderDescending
Private LvwSortTypeBinary, LvwSortTypeText, LvwSortTypeNumeric, LvwSortTypeCurrency, LvwSortTypeDate, LvwSortTypeLogical
Private LvwPictureAlignmentTopLeft, LvwPictureAlignmentTopRight, LvwPictureAlignmentBottomLeft, LvwPictureAlignmentBottomRight, LvwPictureAlignmentCenter, LvwPictureAlignmentTile
Private LvwGroupHeaderAlignmentLeft, LvwGroupHeaderAlignmentRight, LvwGroupHeaderAlignmentCenter
Private LvwGroupFooterAlignmentLeft, LvwGroupFooterAlignmentRight, LvwGroupFooterAlignmentCenter
Private LvwVisualThemeStandard, LvwVisualThemeExplorer
Private LvwVirtualPropertyText, LvwVirtualPropertyIcon, LvwVirtualPropertyIndentation, LvwVirtualPropertyToolTipText, LvwVirtualPropertyBold, LvwVirtualPropertyForeColor, LvwVirtualPropertyChecked
Private LvwFindDirectionUndefined, LvwFindDirectionPrior, LvwFindDirectionNext, LvwFindDirectionEnd, LvwFindDirectionHome, LvwFindDirectionLeft, LvwFindDirectionUp, LvwFindDirectionRight, LvwFindDirectionDown
#End If
Public Enum LvwViewConstants
LvwViewIcon = 0
LvwViewSmallIcon = 1
LvwViewList = 2
LvwViewReport = 3
LvwViewTile = 4
End Enum
Public Enum LvwArrangeConstants
LvwArrangeNone = 0
LvwArrangeAutoLeft = 1
LvwArrangeAutoTop = 2
LvwArrangeLeft = 3
LvwArrangeTop = 4
End Enum
Public Enum LvwColumnHeaderAlignmentConstants
LvwColumnHeaderAlignmentLeft = 0
LvwColumnHeaderAlignmentRight = 1
LvwColumnHeaderAlignmentCenter = 2
End Enum
Public Enum LvwColumnHeaderSortArrowConstants
LvwColumnHeaderSortArrowNone = 0
LvwColumnHeaderSortArrowDown = 1
LvwColumnHeaderSortArrowUp = 2
End Enum
Public Enum LvwColumnHeaderAutoSizeConstants
LvwColumnHeaderAutoSizeToItems = 0
LvwColumnHeaderAutoSizeToHeader = 1
End Enum
Public Enum LvwColumnHeaderFilterTypeConstants
LvwColumnHeaderFilterTypeText = 0
LvwColumnHeaderFilterTypeNumber = 1
End Enum
Public Enum LvwLabelEditConstants
LvwLabelEditAutomatic = 0
LvwLabelEditManual = 1
LvwLabelEditDisabled = 2
End Enum
Public Enum LvwSortOrderConstants
LvwSortOrderAscending = 0
LvwSortOrderDescending = 1
End Enum
Public Enum LvwSortTypeConstants
LvwSortTypeBinary = 0
LvwSortTypeText = 1
LvwSortTypeNumeric = 2
LvwSortTypeCurrency = 3
LvwSortTypeDate = 4
LvwSortTypeLogical = 5
End Enum
Public Enum LvwPictureAlignmentConstants
LvwPictureAlignmentTopLeft = 0
LvwPictureAlignmentTopRight = 1
LvwPictureAlignmentBottomLeft = 2
LvwPictureAlignmentBottomRight = 3
LvwPictureAlignmentCenter = 4
LvwPictureAlignmentTile = 5
End Enum
Public Enum LvwGroupHeaderAlignmentConstants
LvwGroupHeaderAlignmentLeft = 0
LvwGroupHeaderAlignmentRight = 1
LvwGroupHeaderAlignmentCenter = 2
End Enum
Public Enum LvwGroupFooterAlignmentConstants
LvwGroupFooterAlignmentLeft = 0
LvwGroupFooterAlignmentRight = 1
LvwGroupFooterAlignmentCenter = 2
End Enum
Public Enum LvwVisualThemeConstants
LvwVisualThemeStandard = 0
LvwVisualThemeExplorer = 1
End Enum
Public Enum LvwVirtualPropertyConstants
LvwVirtualPropertyText = 1
LvwVirtualPropertyIcon = 2
LvwVirtualPropertyIndentation = 4
LvwVirtualPropertyToolTipText = 8
LvwVirtualPropertyBold = 16
LvwVirtualPropertyForeColor = 32
LvwVirtualPropertyChecked = 64
End Enum
Public Enum LvwFindDirectionConstants
LvwFindDirectionUndefined = 0
LvwFindDirectionPrior = vbKeyPageUp
LvwFindDirectionNext = vbKeyPageDown
LvwFindDirectionEnd = vbKeyEnd
LvwFindDirectionHome = vbKeyHome
LvwFindDirectionLeft = vbKeyLeft
LvwFindDirectionUp = vbKeyUp
LvwFindDirectionRight = vbKeyRight
LvwFindDirectionDown = vbKeyDown
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
Private Type SIZEAPI
CX As Long
CY As Long
End Type
Private Type WINDOWPOS
hWnd As Long
hWndInsertAfter As Long
X As Long
Y As Long
CX As Long
CY As Long
Flags As Long
End Type
Private Type LVITEM
Mask As Long
iItem As Long
iSubItem As Long
State As Long
StateMask As Long
pszText As Long
cchTextMax As Long
iImage As Long
lParam As Long
iIndent As Long
End Type
Private Type LVITEM_V60
LVI As LVITEM
iGroupId As Long
cColumns As Long
puColumns As Long
End Type
Private Type LVTILEINFO
cbSize As Long
iItem As Long
cColumns As Long
puColumns As Long
End Type
Private Type LVTILEVIEWINFO
cbSize As Long
dwMask As Long
dwFlags As Long
SizeTile As SIZEAPI
cLines As Long
RCLabelMargin As RECT
End Type
Private Type LVFINDINFO
Flags As Long
psz As Long
lParam As Long
PT As POINTAPI
VKDirection As Long
End Type
Private Type LVCOLUMN
Mask As Long
fmt As Long
CX As Long
pszText As Long
cchTextMax As Long
iSubItem As Long
iImage As Long
iOrder As Long
End Type
Private Type LVHITTESTINFO
PT As POINTAPI
Flags As Long
iItem As Long
iSubItem As Long
End Type
Private Type LVINSERTMARK
cbSize As Long
dwFlags As Long
iItem As Long
dwReserved As Long
End Type
Private Type LVBKIMAGE
ulFlags As Long
hBmp As Long
pszImage As String
cchImageMax As Long
XOffsetPercent As Long
YOffsetPercent As Long
End Type
Private Type LVGROUP
cbSize As Long
Mask As Long
pszHeader As Long
cchHeader As Long
pszFooter As Long
cchFooter As Long
iGroupId As Long
StateMask As Long
State As Long
uAlign As Long
End Type
Private Type LVGROUP_V61
LVG As LVGROUP
pszSubtitle As Long
cchSubtitle As Long
pszTask As Long
cchTask As Long
pszDescriptionTop As Long
cchDescriptionTop As Long
pszDescriptionBottom As Long
cchDescriptionBottom As Long
iTitleImage As Long
iExtendedImage As Long
iFirstItem As Long
cItems As Long
pszSubsetTitle As Long
cchSubsetTitle As Long
End Type
Private Type LVINSERTGROUPSORTED
pfnGroupCompare As Long
pvData As ISubclass
LVG As LVGROUP
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
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM + 1)
Private Const CDDS_SUBITEM As Long = &H20000
Private Const CDIS_HOT As Long = &H40
Private Const CDRF_NEWFONT As Long = &H2
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20
Private Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20
Private Type NMCUSTOMDRAW
hdr As NMHDR
dwDrawStage As Long
hDC As Long
RC As RECT
dwItemSpec As Long
uItemState As Long
lItemlParam As Long
End Type
Private Type NMLVCUSTOMDRAW
NMCD As NMCUSTOMDRAW
ClrText As Long
ClrTextBk As Long
iSubItem As Long
End Type
Private Type NMLISTVIEW
hdr As NMHDR
iItem As Long
iSubItem As Long
uNewState As Long
uOldState As Long
uChanged As Long
PTAction As POINTAPI
lParam As Long
End Type
Private Type NMITEMACTIVATE
hdr As NMHDR
iItem As Long
iSubItem As Long
uNewState As Long
uOldState As Long
uChanged As Long
PTAction As POINTAPI
lParam As Long
uKeyFlags As Long
End Type
Private Type NMLVGETINFOTIP
hdr As NMHDR
dwFlags As Long
pszText As Long
cchTextMax As Long
iItem As Long
iSubItem As Long
lParam As Long
End Type
Private Type NMLVDISPINFO
hdr As NMHDR
Item As LVITEM
End Type
Private Const L_MAX_URL_LENGTH As Long = 2084
Private Type NMLVEMPTYMARKUP
hdr As NMHDR
dwFlags As Long
szMarkup(0 To ((L_MAX_URL_LENGTH * 2) - 1)) As Byte
End Type
Private Type NMLVFINDITEM
hdr As NMHDR
iStart As Long
LVFI As LVFINDINFO
End Type
Private Type NMLVCACHEHINT
hdr As NMHDR
iFrom As Long
iTo As Long
End Type
Private Type NMLVODSTATECHANGE
hdr As NMHDR
iFrom As Long
iTo As Long
uNewState As Long
uOldState As Long
End Type
Private Type NMLVSCROLL
hdr As NMHDR
DX As Long
DY As Long
End Type
Private Type NMLVGROUP
hdr As NMHDR
iGroupId As Long
uNewState As Long
uOldState As Long
End Type
Private Const MAX_LINKID_TEXT As Long = 48
Private Type LITEM
Mask As Long
iLink As Long
State As Long
StateMask As Long
szID(0 To ((MAX_LINKID_TEXT * 2) - 1)) As Byte
szURL(0 To ((L_MAX_URL_LENGTH * 2) - 1)) As Byte
End Type
Private Type NMLVLINK
hdr As NMHDR
Item As LITEM
iItem As Long
iGroupId As Long
End Type
Private Type NMHEADER
hdr As NMHDR
iItem As Long
iButton As Long
lPtrHDItem As Long
End Type
Private Type NMHDFILTERBTNCLICK
hdr As NMHDR
iItem As Long
RC As RECT
End Type
Private Type NMTTDISPINFO
hdr As NMHDR
lpszText As Long
szText(0 To ((80 * 2) - 1)) As Byte
hInst As Long
uFlags As Long
End Type
Private Type HDITEM
Mask As Long
CXY As Long
pszText As Long
hBm As Long
cchTextMax As Long
fmt As Long
lParam As Long
iImage As Long
iOrder As Long
FilterType As Long
pvFilter As Long
End Type
Private Type HDHITTESTINFO
PT As POINTAPI
Flags As Long
iItem As Long
End Type
Private Type HDLAYOUT
lpRC As Long
lpWPOS As Long
End Type
Private Type HDTEXTFILTER
pszText As Long
cchTextMax As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event BeforeScroll(ByVal DeltaX As Single, ByVal DeltaY As Single)
Attribute BeforeScroll.VB_Description = "Occurs when the control is about to be scrolled. Requires comctl32.dll version 6.0 or higher."
Public Event AfterScroll(ByVal DeltaX As Single, ByVal DeltaY As Single)
Attribute AfterScroll.VB_Description = "Occurs when the control has been scrolled. Requires comctl32.dll version 6.0 or higher."
Public Event ContextMenu(ByVal X As Single, ByVal Y As Single)
Attribute ContextMenu.VB_Description = "Occurs when the user clicked the right mouse button or types SHIFT + F10."
Public Event ItemClick(ByVal Item As LvwListItem, ByVal Button As Integer)
Attribute ItemClick.VB_Description = "Occurs when a list item is clicked."
Public Event ItemDblClick(ByVal Item As LvwListItem, ByVal Button As Integer)
Attribute ItemDblClick.VB_Description = "Occurs when a list item is double clicked."
Public Event ItemFocus(ByVal Item As LvwListItem)
Attribute ItemFocus.VB_Description = "Occurs when a list item is focused."
Public Event ItemActivate(ByVal Item As LvwListItem, ByVal SubItemIndex As Long, ByVal Shift As Integer)
Attribute ItemActivate.VB_Description = "Occurs when a list item is activated."
Public Event ItemSelect(ByVal Item As LvwListItem, ByVal Selected As Boolean)
Attribute ItemSelect.VB_Description = "Occurs when a list item is selected."
Public Event ItemCheck(ByVal Item As LvwListItem, ByVal Checked As Boolean)
Attribute ItemCheck.VB_Description = "Occurs when a list item is checked."
Public Event ItemDrag(ByVal Item As LvwListItem, ByVal Button As Integer)
Attribute ItemDrag.VB_Description = "Occurs when a list item initiate a drag-and-drop operation."
Public Event ItemBkColor(ByVal Item As LvwListItem, ByRef RGBColor As Long)
Attribute ItemBkColor.VB_Description = "Occurs when a list item is about to draw the background in 'report' view. This is a request to provide an alternative back color. The back color is passed in an RGB format."
Public Event GetVirtualItem(ByVal ItemIndex As Long, ByVal SubItemIndex As Long, ByVal VirtualProperty As LvwVirtualPropertyConstants, ByRef Value As Variant)
Attribute GetVirtualItem.VB_Description = "Occurs when the list view is in virtual mode and requests for an item or sub item property."
Public Event FindVirtualItem(ByVal StartIndex As Long, ByVal SearchText As String, ByVal Partial As Boolean, ByVal Wrap As Boolean, ByRef FoundIndex As Long)
Attribute FindVirtualItem.VB_Description = "Occurs when the list view is in virtual mode and needs to find a particular item."
Public Event CacheVirtualItems(ByVal FromIndex As Long, ByVal ToIndex As Long)
Attribute CacheVirtualItems.VB_Description = "Occurs when the list view is in virtual mode and the contents of its display area have changed. It contains information about the range of items to be cached."
Public Event BeforeLabelEdit(ByRef Cancel As Boolean)
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected list item."
Public Event AfterLabelEdit(ByRef Cancel As Boolean, ByRef NewString As String)
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected list item."
Public Event ColumnClick(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnClick.VB_Description = "Occurs when a column header in a list view is clicked."
Public Event ColumnDblClick(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnDblClick.VB_Description = "Occurs when a column header in a list view is double-clicked."
Public Event ColumnBeforeResize(ByVal ColumnHeader As LvwColumnHeader, ByRef Cancel As Boolean)
Attribute ColumnBeforeResize.VB_Description = "Occurs when the user has begun dragging a divider on one column header."
Public Event ColumnAfterResize(ByVal ColumnHeader As LvwColumnHeader, ByRef NewWidth As Single)
Attribute ColumnAfterResize.VB_Description = "Occurs when the user has finished dragging a divider on one column header."
Public Event ColumnDividerDblClick(ByVal ColumnHeader As LvwColumnHeader, ByRef Cancel As Boolean)
Attribute ColumnDividerDblClick.VB_Description = "Occurs when the user double-clicked the divider on one column header."
Public Event ColumnBeforeDrag(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnBeforeDrag.VB_Description = "Occurs when a drag operation has begun on one column header."
Public Event ColumnAfterDrag(ByVal ColumnHeader As LvwColumnHeader, ByVal NewPosition As Long, ByRef Cancel As Boolean)
Attribute ColumnAfterDrag.VB_Description = "Occurs when a drag operation has ended on one column header."
Public Event ColumnDropDown(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnDropDown.VB_Description = "Occurs when the drop-down arrow on the split button of a column header is clicked. Requires comctl32.dll version 6.1 or higher."
Public Event ColumnCheck(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnCheck.VB_Description = "Occurs when a column header is checked. Requires comctl32.dll version 6.1 or higher."
Public Event ColumnChevronPushed(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnChevronPushed.VB_Description = "Occurs when a chevron button of a column header is pushed. Requires comctl32.dll version 6.1 or higher."
Public Event ColumnFilterChanged(ByVal ColumnHeader As LvwColumnHeader)
Attribute ColumnFilterChanged.VB_Description = "Occurs when a filter of a column header has been changed."
Public Event ColumnFilterButtonClick(ByVal ColumnHeader As LvwColumnHeader, ByRef RaiseFilterChanged As Boolean, ByVal ButtonLeft As Long, ByVal ButtonTop As Long, ByVal ButtonRight As Long, ByVal ButtonBottom As Long)
Attribute ColumnFilterButtonClick.VB_Description = "Occurs when a filter button of a column header is clicked."
Public Event BeforeFilterEdit(ByVal ColumnHeader As LvwColumnHeader, ByVal hWndFilterEdit As Long)
Attribute BeforeFilterEdit.VB_Description = "Occurs when a user attempts to edit the filter of the corresponding column header."
Public Event AfterFilterEdit(ByVal ColumnHeader As LvwColumnHeader)
Attribute AfterFilterEdit.VB_Description = "Occurs after a user edits the filter of the corresponding column header."
Public Event GetEmptyMarkup(ByRef Text As String, ByRef Center As Boolean)
Attribute GetEmptyMarkup.VB_Description = "Occurs when the list view has no list items. This is a request to provide a markup text. Requires comctl32.dll version 6.1 or higher."
Public Event GroupCollapsedChanged(ByVal Group As LvwGroup)
Attribute GroupCollapsedChanged.VB_Description = "Occurrs when the group's collapsed state changes. Requires comctl32.dll version 6.1 or higher."
Public Event GroupSelectedChanged(ByVal Group As LvwGroup)
Attribute GroupSelectedChanged.VB_Description = "Occurrs when the group's selected state changes. Requires comctl32.dll version 6.1 or higher."
Public Event GroupLinkClick(ByVal Group As LvwGroup)
Attribute GroupLinkClick.VB_Description = "Occurs when a link in a group is clicked. Requires comctl32.dll version 6.1 or higher."
Public Event BeginMarqueeSelection(ByRef Cancel As Boolean)
Attribute BeginMarqueeSelection.VB_Description = "Occurs when a bounding box (marquee) selection has begun. Only applicable if the multi select property is set to true."
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
Private Declare Function StrCmpLogical Lib "shlwapi" Alias "StrCmpLogicalW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function SetWindowTheme Lib "uxtheme" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, ByRef CX As Long, ByRef CY As Long) As Long
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
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Const ICC_LISTVIEW_CLASSES As Long = &H1
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const CF_UNICODETEXT As Long = 13
Private Const LPSTR_TEXTCALLBACK As Long = (-1)
Private Const MAXINT_4 As Long = 2147483647
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_TOPMOST As Long = &H8
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114
Private Const SB_LINELEFT As Long = 0, SB_LINERIGHT As Long = 1
Private Const SB_LINEUP As Long = 0, SB_LINEDOWN As Long = 1
Private Const SW_HIDE As Long = &H0
Private Const SW_SHOW As Long = &H5
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_NOTIFYFORMAT As Long = &H55
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
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SETFONT As Long = &H30
Private Const WM_SIZE As Long = &H5
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const COLOR_HOTLIGHT As Long = 26
Private Const CLR_NONE As Long = &HFFFFFFFF
Private Const CLR_DEFAULT As Long = &HFF000000
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETVERSION As Long = (CCM_FIRST + 7)
Private Const WM_USER As Long = &H400
Private Const UM_ENDFILTEREDIT As Long = (WM_USER + 400)
Private Const UM_BUTTONDOWN As Long = (WM_USER + 500)
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETBKCOLOR As Long = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR As Long = (LVM_FIRST + 1)
Private Const LVM_GETIMAGELIST As Long = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Private Const LVM_GETITEMA As Long = (LVM_FIRST + 5)
Private Const LVM_GETITEMW As Long = (LVM_FIRST + 75)
Private Const LVM_GETITEM As Long = LVM_GETITEMW
Private Const LVM_SETITEMA As Long = (LVM_FIRST + 6)
Private Const LVM_SETITEMW As Long = (LVM_FIRST + 76)
Private Const LVM_SETITEM As Long = LVM_SETITEMW
Private Const LVM_INSERTITEMA As Long = (LVM_FIRST + 7)
Private Const LVM_INSERTITEMW As Long = (LVM_FIRST + 77)
Private Const LVM_INSERTITEM As Long = LVM_INSERTITEMW
Private Const LVM_DELETEITEM As Long = (LVM_FIRST + 8)
Private Const LVM_DELETEALLITEMS As Long = (LVM_FIRST + 9)
Private Const LVM_GETCALLBACKMASK As Long = (LVM_FIRST + 10)
Private Const LVM_SETCALLBACKMASK As Long = (LVM_FIRST + 11)
Private Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)
Private Const LVM_FINDITEMA As Long = (LVM_FIRST + 13)
Private Const LVM_FINDITEMW As Long = (LVM_FIRST + 83)
Private Const LVM_FINDITEM As Long = LVM_FINDITEMW
Private Const LVM_RESETEMPTYTEXT As Long = (LVM_FIRST + 84) ' Undocumented
Private Const LVM_GETITEMRECT As Long = (LVM_FIRST + 14)
Private Const LVM_SETITEMPOSITION As Long = (LVM_FIRST + 15) ' 16 bit
Private Const LVM_GETITEMPOSITION As Long = (LVM_FIRST + 16)
Private Const LVM_GETSTRINGWIDTHA As Long = (LVM_FIRST + 17)
Private Const LVM_GETSTRINGWIDTHW As Long = (LVM_FIRST + 87)
Private Const LVM_GETSTRINGWIDTH As Long = LVM_GETSTRINGWIDTHW
Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE As Long = (LVM_FIRST + 19)
Private Const LVM_SCROLL As Long = (LVM_FIRST + 20)
Private Const LVM_REDRAWITEMS As Long = (LVM_FIRST + 21)
Private Const LVM_ARRANGE As Long = (LVM_FIRST + 22)
Private Const LVM_EDITLABELA As Long = (LVM_FIRST + 23)
Private Const LVM_EDITLABELW As Long = (LVM_FIRST + 118)
Private Const LVM_EDITLABEL As Long = LVM_EDITLABELW
Private Const LVM_GETEDITCONTROL As Long = (LVM_FIRST + 24)
Private Const LVM_GETCOLUMNA As Long = (LVM_FIRST + 25)
Private Const LVM_GETCOLUMNW As Long = (LVM_FIRST + 95)
Private Const LVM_GETCOLUMN As Long = LVM_GETCOLUMNW
Private Const LVM_SETCOLUMNA As Long = (LVM_FIRST + 26)
Private Const LVM_SETCOLUMNW As Long = (LVM_FIRST + 96)
Private Const LVM_SETCOLUMN As Long = LVM_SETCOLUMNW
Private Const LVM_INSERTCOLUMNA As Long = (LVM_FIRST + 27)
Private Const LVM_INSERTCOLUMNW As Long = (LVM_FIRST + 97)
Private Const LVM_INSERTCOLUMN As Long = LVM_INSERTCOLUMNW
Private Const LVM_DELETECOLUMN As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Private Const LVM_CREATEDRAGIMAGE As Long = (LVM_FIRST + 33)
Private Const LVM_GETVIEWRECT As Long = (LVM_FIRST + 34)
Private Const LVM_GETTEXTCOLOR As Long = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR As Long = (LVM_FIRST + 36)
Private Const LVM_GETTEXTBKCOLOR As Long = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR As Long = (LVM_FIRST + 38)
Private Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE As Long = (LVM_FIRST + 40)
Private Const LVM_GETORIGIN As Long = (LVM_FIRST + 41)
Private Const LVM_UPDATE As Long = (LVM_FIRST + 42)
Private Const LVM_SETITEMSTATE As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
Private Const LVM_GETITEMTEXTW As Long = (LVM_FIRST + 115)
Private Const LVM_GETITEMTEXT As Long = LVM_GETITEMTEXTW
Private Const LVM_SETITEMTEXTA As Long = (LVM_FIRST + 46)
Private Const LVM_SETITEMTEXTW As Long = (LVM_FIRST + 116)
Private Const LVM_SETITEMTEXT As Long = LVM_SETITEMTEXTW
Private Const LVM_SETITEMCOUNT As Long = (LVM_FIRST + 47)
Private Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
Private Const LVM_SETITEMPOSITION32 As Long = (LVM_FIRST + 49)
Private Const LVM_GETSELECTEDCOUNT As Long = (LVM_FIRST + 50)
Private Const LVM_GETITEMSPACING As Long = (LVM_FIRST + 51)
Private Const LVM_GETISEARCHSTRINGA As Long = (LVM_FIRST + 52)
Private Const LVM_GETISEARCHSTRINGW As Long = (LVM_FIRST + 117)
Private Const LVM_GETISEARCHSTRING As Long = LVM_GETISEARCHSTRINGW
Private Const LVM_SETICONSPACING As Long = (LVM_FIRST + 53)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Private Const LVM_GETSUBITEMRECT As Long = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Private Const LVM_SETCOLUMNORDERARRAY As Long = (LVM_FIRST + 58)
Private Const LVM_GETCOLUMNORDERARRAY As Long = (LVM_FIRST + 59)
Private Const LVM_SETHOTITEM As Long = (LVM_FIRST + 60)
Private Const LVM_GETHOTITEM As Long = (LVM_FIRST + 61)
Private Const LVM_SETHOTCURSOR As Long = (LVM_FIRST + 62)
Private Const LVM_GETHOTCURSOR As Long = (LVM_FIRST + 63)
Private Const LVM_APPROXIMATEVIEWRECT As Long = (LVM_FIRST + 64)
Private Const LVM_SETWORKAREAS As Long = (LVM_FIRST + 65)
Private Const LVM_GETSELECTIONMARK As Long = (LVM_FIRST + 66)
Private Const LVM_SETSELECTIONMARK As Long = (LVM_FIRST + 67)
Private Const LVM_SETBKIMAGEA As Long = (LVM_FIRST + 68)
Private Const LVM_SETBKIMAGEW As Long = (LVM_FIRST + 138)
Private Const LVM_SETBKIMAGE As Long = LVM_SETBKIMAGEW
Private Const LVM_GETBKIMAGEA As Long = (LVM_FIRST + 69)
Private Const LVM_GETBKIMAGEW As Long = (LVM_FIRST + 139)
Private Const LVM_GETBKIMAGE As Long = LVM_GETBKIMAGEW
Private Const LVM_GETWORKAREAS As Long = (LVM_FIRST + 70)
Private Const LVM_SETHOVERTIME As Long = (LVM_FIRST + 71)
Private Const LVM_GETHOVERTIME As Long = (LVM_FIRST + 72)
Private Const LVM_GETNUMBEROFWORKAREAS As Long = (LVM_FIRST + 73)
Private Const LVM_SETTOOLTIPS As Long = (LVM_FIRST + 74)
Private Const LVM_GETTOOLTIPS As Long = (LVM_FIRST + 78)
Private Const LVM_GETHOTLIGHTCOLOR As Long = (LVM_FIRST + 79) ' Undocumented
Private Const LVM_SETHOTLIGHTCOLOR As Long = (LVM_FIRST + 80) ' Undocumented
Private Const LVM_SORTITEMSEX As Long = (LVM_FIRST + 81)
Private Const LVM_GETGROUPSTATE As Long = (LVM_FIRST + 92)
Private Const LVM_GETFOCUSEDGROUP As Long = (LVM_FIRST + 93)
Private Const LVM_GETGROUPRECT As Long = (LVM_FIRST + 98)
Private Const LVM_SETSELECTEDCOLUMN As Long = (LVM_FIRST + 140)
Private Const LVM_SETVIEW As Long = (LVM_FIRST + 142)
Private Const LVM_GETVIEW As Long = (LVM_FIRST + 143)
Private Const LVM_INSERTGROUP As Long = (LVM_FIRST + 145)
Private Const LVM_SETGROUPINFO As Long = (LVM_FIRST + 147)
Private Const LVM_GETGROUPINFO As Long = (LVM_FIRST + 149)
Private Const LVM_REMOVEGROUP As Long = (LVM_FIRST + 150)
Private Const LVM_GETGROUPCOUNT As Long = (LVM_FIRST + 152)
Private Const LVM_GETGROUPINFOBYINDEX As Long = (LVM_FIRST + 153)
Private Const LVM_ENABLEGROUPVIEW As Long = (LVM_FIRST + 157)
Private Const LVM_SORTGROUPS As Long = (LVM_FIRST + 158)
Private Const LVM_INSERTGROUPSORTED As Long = (LVM_FIRST + 159)
Private Const LVM_REMOVEALLGROUPS As Long = (LVM_FIRST + 160)
Private Const LVM_HASGROUP As Long = (LVM_FIRST + 161)
Private Const LVM_SETTILEVIEWINFO As Long = (LVM_FIRST + 162)
Private Const LVM_GETTILEVIEWINFO As Long = (LVM_FIRST + 163)
Private Const LVM_SETTILEINFO As Long = (LVM_FIRST + 164)
Private Const LVM_GETTILEINFO As Long = (LVM_FIRST + 165)
Private Const LVM_SETINSERTMARK As Long = (LVM_FIRST + 166)
Private Const LVM_GETINSERTMARK As Long = (LVM_FIRST + 167)
Private Const LVM_INSERTMARKHITTEST As Long = (LVM_FIRST + 168)
Private Const LVM_GETINSERTMARKRECT As Long = (LVM_FIRST + 169)
Private Const LVM_SETINSERTMARKCOLOR As Long = (LVM_FIRST + 170)
Private Const LVM_GETINSERTMARKCOLOR As Long = (LVM_FIRST + 171)
Private Const LVM_SETINFOTIP As Long = (LVM_FIRST + 173)
Private Const LVM_GETSELECTEDCOLUMN As Long = (LVM_FIRST + 174)
Private Const LVM_ISGROUPVIEWENABLED As Long = (LVM_FIRST + 175)
Private Const LVM_CANCELEDITLABEL As Long = (LVM_FIRST + 179)
Private Const LVM_ISITEMVISIBLE As Long = (LVM_FIRST + 182)
Private Const LVM_SETGROUPSUBSETCOUNT As Long = (LVM_FIRST + 190) ' Undocumented
Private Const LVM_GETGROUPSUBSETCOUNT As Long = (LVM_FIRST + 191) ' Undocumented
Private Const LVN_FIRST As Long = (-100)
Private Const LVN_ITEMCHANGING As Long = (LVN_FIRST - 0)
Private Const LVN_ITEMCHANGED As Long = (LVN_FIRST - 1)
Private Const LVN_INSERTITEM As Long = (LVN_FIRST - 2)
Private Const LVN_DELETEITEM As Long = (LVN_FIRST - 3)
Private Const LVN_DELETEALLITEMS As Long = (LVN_FIRST - 4)
Private Const LVN_BEGINLABELEDITA As Long = (LVN_FIRST - 5)
Private Const LVN_BEGINLABELEDITW As Long = (LVN_FIRST - 75)
Private Const LVN_BEGINLABELEDIT As Long = LVN_BEGINLABELEDITW
Private Const LVN_ENDLABELEDITA As Long = (LVN_FIRST - 6)
Private Const LVN_ENDLABELEDITW As Long = (LVN_FIRST - 76)
Private Const LVN_ENDLABELEDIT As Long = LVN_ENDLABELEDITW
Private Const LVN_COLUMNCLICK As Long = (LVN_FIRST - 8)
Private Const LVN_BEGINDRAG As Long = (LVN_FIRST - 9)
Private Const LVN_BEGINRDRAG As Long = (LVN_FIRST - 11)
Private Const LVN_ODCACHEHINT As Long = (LVN_FIRST - 13)
Private Const LVN_ITEMACTIVATE As Long = (LVN_FIRST - 14)
Private Const LVN_ODSTATECHANGED As Long = (LVN_FIRST - 15)
Private Const LVN_HOTTRACK As Long = (LVN_FIRST - 21)
Private Const LVN_GETDISPINFOA As Long = (LVN_FIRST - 50)
Private Const LVN_GETDISPINFOW As Long = (LVN_FIRST - 77)
Private Const LVN_GETDISPINFO As Long = LVN_GETDISPINFOW
Private Const LVN_SETDISPINFOA As Long = (LVN_FIRST - 51)
Private Const LVN_SETDISPINFOW As Long = (LVN_FIRST - 78)
Private Const LVN_SETDISPINFO As Long = LVN_SETDISPINFOW
Private Const LVN_ODFINDITEMA As Long = (LVN_FIRST - 52)
Private Const LVN_ODFINDITEMW As Long = (LVN_FIRST - 79)
Private Const LVN_ODFINDITEM As Long = LVN_ODFINDITEMW
Private Const LVN_MARQUEEBEGIN As Long = (LVN_FIRST - 56)
Private Const LVN_GETINFOTIPA As Long = (LVN_FIRST - 57)
Private Const LVN_GETINFOTIPW As Long = (LVN_FIRST - 58)
Private Const LVN_GETINFOTIP As Long = LVN_GETINFOTIPW
Private Const LVN_INCREMENTALSEARCHA As Long = (LVN_FIRST - 62)
Private Const LVN_INCREMENTALSEARCHW As Long = (LVN_FIRST - 63)
Private Const LVN_INCREMENTALSEARCH As Long = LVN_INCREMENTALSEARCHW
Private Const LVN_COLUMNOVERFLOWCLICK As Long = (LVN_FIRST - 66)
Private Const LVN_BEGINSCROLL As Long = (LVN_FIRST - 80)
Private Const LVN_ENDSCROLL As Long = (LVN_FIRST - 81)
Private Const LVN_LINKCLICK As Long = (LVN_FIRST - 84)
Private Const LVN_GETEMPTYMARKUP As Long = (LVN_FIRST - 87)
Private Const LVN_GROUPCHANGED As Long = (LVN_FIRST - 88) ' Undocumented
Private Const LVA_DEFAULT As Long = &H0
Private Const LVNI_ALL As Long = &H0
Private Const LVNI_FOCUSED As Long = &H1
Private Const LVNI_SELECTED As Long = &H2
Private Const LVNI_CUT As Long = &H4
Private Const LVNI_DROPHILITED As Long = &H8
Private Const LVNI_VISIBLEORDER As Long = &H10
Private Const LVNI_VISIBLEONLY As Long = &H40
Private Const LVNI_ABOVE As Long = &H100
Private Const LVNI_BELOW As Long = &H200
Private Const LVNI_TOLEFT As Long = &H400
Private Const LVNI_TORIGHT As Long = &H800
Private Const LVIF_TEXT As Long = &H1
Private Const LVIF_IMAGE As Long = &H2
Private Const LVIF_PARAM As Long = &H4
Private Const LVIF_STATE As Long = &H8
Private Const LVIF_INDENT As Long = &H10
Private Const LVIF_GROUPID As Long = &H100
Private Const LVIF_COLUMNS As Long = &H200
Private Const LVIF_NORECOMPUTE As Long = &H800
Private Const LVIR_BOUNDS As Long = 0
Private Const LVIR_ICON As Long = 1
Private Const LVIR_LABEL As Long = 2
Private Const LVIR_SELECTBOUNDS As Long = 3
Private Const LVIS_FOCUSED As Long = &H1
Private Const LVIS_SELECTED As Long = &H2
Private Const LVIS_CUT As Long = &H4
Private Const LVIS_DROPHILITED As Long = &H8
Private Const LVIS_ACTIVATING As Long = &H20 ' Unsupported
Private Const LVIS_OVERLAYMASK As Long = &HF00
Private Const LVIS_STATEIMAGEMASK As Long = &HF000&
Private Const LVFI_PARAM As Long = &H1
Private Const LVFI_STRING As Long = &H2
Private Const LVFI_PARTIAL As Long = &H8
Private Const LVFI_WRAP As Long = &H20
Private Const LVFI_NEARESTXY As Long = &H40
Private Const LVKF_ALT As Long = &H1
Private Const LVKF_CONTROL As Long = &H2
Private Const LVKF_SHIFT As Long = &H4
Private Const LVBKIF_SOURCE_NONE As Long = &H0
Private Const LVBKIF_SOURCE_HBITMAP As Long = &H1
Private Const LVBKIF_SOURCE_URL As Long = &H2
Private Const LVBKIF_STYLE_NORMAL As Long = &H0
Private Const LVBKIF_STYLE_TILE As Long = &H10
Private Const LVBKIF_TYPE_WATERMARK As Long = &H10000000
Private Const LVBKIF_FLAG_TILEOFFSET As Long = &H100
Private Const LVBKIF_FLAG_ALPHABLEND As Long = &H20000000
Private Const LVGIT_UNFOLDED As Long = &H1
Private Const LVSIL_NORMAL As Long = 0
Private Const LVSIL_SMALL As Long = 1
Private Const LVSIL_STATE As Long = 2
Private Const LVSIL_GROUPHEADER As Long = 3
Private Const LVHT_NOWHERE As Long = &H1
Private Const LVHT_ONITEMICON As Long = &H2
Private Const LVHT_ONITEMLABEL As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
Private Const LVHT_ABOVE As Long = &H8
Private Const LVHT_BELOW As Long = &H10
Private Const LVHT_TORIGHT As Long = &H20
Private Const LVHT_TOLEFT As Long = &H40
Private Const LV_MAX_WORKAREAS As Long = 16
Private Const EMF_CENTERED As Long = 1
Private Const HDM_FIRST As Long = &H1200
Private Const HDM_GETITEMA As Long = (HDM_FIRST + 3)
Private Const HDM_GETITEMW As Long = (HDM_FIRST + 11)
Private Const HDM_GETITEM As Long = HDM_GETITEMW
Private Const HDM_SETITEMA As Long = (HDM_FIRST + 4)
Private Const HDM_SETITEMW As Long = (HDM_FIRST + 12)
Private Const HDM_SETITEM As Long = HDM_SETITEMW
Private Const HDM_LAYOUT As Long = (HDM_FIRST + 5)
Private Const HDM_HITTEST As Long = (HDM_FIRST + 6)
Private Const HDM_SETIMAGELIST As Long = (HDM_FIRST + 8)
Private Const HDM_GETIMAGELIST As Long = (HDM_FIRST + 9)
Private Const HDM_ORDERTOINDEX As Long = (HDM_FIRST + 15)
Private Const HDM_SETFILTERCHANGETIMEOUT As Long = (HDM_FIRST + 22)
Private Const HDM_EDITFILTER As Long = (HDM_FIRST + 23)
Private Const HDM_CLEARFILTER As Long = (HDM_FIRST + 24)
Private Const HDM_GETFOCUSEDITEM As Long = (HDM_FIRST + 27)
Private Const HDSIL_NORMAL As Long = 0
Private Const HDSIL_STATE As Long = 1
Private Const HHT_ONDIVIDER As Long = &H4
Private Const HHT_ONDIVOPEN As Long = &H8
Private Const HHT_ONFILTER As Long = &H10
Private Const HHT_ONFILTERBUTTON As Long = &H20
Private Const HHT_ONDROPDOWN As Long = &H2000
Private Const HDI_WIDTH As Long = &H1
Private Const HDI_FORMAT As Long = &H4
Private Const HDI_ORDER As Long = &H80
Private Const HDI_FILTER As Long = &H100
Private Const HDFT_ISSTRING As Long = &H0
Private Const HDFT_ISNUMBER As Long = &H1
Private Const HDFT_HASNOVALUE As Long = &H8000&
Private Const HDF_RTLREADING As Long = &H4
Private Const HDF_SORTDOWN As Long = &H200
Private Const HDF_SORTUP As Long = &H400
Private Const HDF_CHECKBOX As Long = &H40
Private Const HDF_CHECKED As Long = &H80
Private Const HDS_BUTTONS As Long = &H2
Private Const HDS_HOTTRACK As Long = &H4
Private Const HDS_CHECKBOXES As Long = &H400
Private Const HDS_FULLDRAG As Long = &H80
Private Const HDS_FILTERBAR As Long = &H100
Private Const HDS_NOSIZING As Long = &H800
Private Const HDS_OVERFLOW As Long = &H1000
Private Const HDN_FIRST As Long = (-300)
Private Const HDN_ITEMDBLCLICKA As Long = (HDN_FIRST - 3)
Private Const HDN_ITEMDBLCLICKW As Long = (HDN_FIRST - 23)
Private Const HDN_ITEMDBLCLICK As Long = HDN_ITEMDBLCLICKW
Private Const HDN_DIVIDERDBLCLICKA As Long = (HDN_FIRST - 5)
Private Const HDN_DIVIDERDBLCLICKW As Long = (HDN_FIRST - 25)
Private Const HDN_DIVIDERDBLCLICK As Long = HDN_DIVIDERDBLCLICKW
Private Const HDN_BEGINTRACKA As Long = (HDN_FIRST - 6)
Private Const HDN_BEGINTRACKW As Long = (HDN_FIRST - 26)
Private Const HDN_BEGINTRACK As Long = HDN_BEGINTRACKW
Private Const HDN_ENDTRACKA As Long = (HDN_FIRST - 7)
Private Const HDN_ENDTRACKW As Long = (HDN_FIRST - 27)
Private Const HDN_ENDTRACK As Long = HDN_ENDTRACKW
Private Const HDN_BEGINDRAG As Long = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG As Long = (HDN_FIRST - 11)
Private Const HDN_FILTERCHANGE As Long = (HDN_FIRST - 12)
Private Const HDN_FILTERBTNCLICK As Long = (HDN_FIRST - 13)
Private Const HDN_BEGINFILTEREDIT As Long = (HDN_FIRST - 14)
Private Const HDN_ENDFILTEREDIT As Long = (HDN_FIRST - 15)
Private Const HDN_ITEMSTATEICONCLICK As Long = (HDN_FIRST - 16)
Private Const HDN_DROPDOWN As Long = (HDN_FIRST - 18)
Private Const TTM_POP As Long = (WM_USER + 28)
Private Const TTM_ADDTOOLA As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW As Long = (WM_USER + 50)
Private Const TTM_ADDTOOL As Long = TTM_ADDTOOLW
Private Const TTM_NEWTOOLRECTA As Long = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW As Long = (WM_USER + 52)
Private Const TTM_NEWTOOLRECT As Long = TTM_NEWTOOLRECTW
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTF_PARSELINKS As Long = &H1000
Private Const TTF_RTLREADING As Long = &H4
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTN_FIRST As Long = (-520)
Private Const TTN_GETDISPINFOA As Long = (TTN_FIRST - 0)
Private Const TTN_GETDISPINFOW As Long = (TTN_FIRST - 10)
Private Const TTN_GETDISPINFO As Long = TTN_GETDISPINFOW
Private Const TTN_SHOW As Long = (TTN_FIRST - 1)
Private Const LVCF_FMT As Long = &H1
Private Const LVCF_WIDTH As Long = &H2
Private Const LVCF_TEXT As Long = &H4
Private Const LVCF_SUBITEM As Long = &H8
Private Const LVCF_IMAGE As Long = &H10
Private Const LVCF_ORDER As Long = &H20
Private Const LVGF_HEADER As Long = &H1
Private Const LVGF_FOOTER As Long = &H2
Private Const LVGF_STATE As Long = &H4
Private Const LVGF_ALIGN As Long = &H8
Private Const LVGF_GROUPID As Long = &H10
Private Const LVGF_SUBTITLE As Long = &H100
Private Const LVGF_TASK As Long = &H200
Private Const LVGF_TITLEIMAGE As Long = &H1000
Private Const LVGF_ITEMS As Long = &H4000
Private Const LVGF_SUBSET As Long = &H8000&
Private Const LVGF_SUBSETITEMS As Long = &H10000
Private Const LVGA_FOOTER_LEFT As Long = &H8
Private Const LVGA_FOOTER_CENTER As Long = &H10
Private Const LVGA_FOOTER_RIGHT As Long = &H20
Private Const LVGA_HEADER_LEFT As Long = &H1
Private Const LVGA_HEADER_CENTER As Long = &H2
Private Const LVGA_HEADER_RIGHT As Long = &H4
Private Const LVGS_NORMAL As Long = &H0
Private Const LVGS_COLLAPSED As Long = &H1
Private Const LVGS_HIDDEN As Long = &H2 ' Malfunction
Private Const LVGS_NOHEADER As Long = &H4
Private Const LVGS_COLLAPSIBLE As Long = &H8
Private Const LVGS_FOCUSED As Long = &H10
Private Const LVGS_SELECTED As Long = &H20
Private Const LVGS_SUBSETED As Long = &H40
Private Const LVGS_SUBSETLINKFOCUSED As Long = &H80
Private Const LVGGR_GROUP As Long = 0
Private Const LVGGR_HEADER As Long = 1
Private Const LVGGR_LABEL As Long = 2
Private Const LVGGR_SUBSETLINK As Long = 3
Private Const LVSCW_AUTOSIZE As Long = (-1)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = (-2)
Private Const LVIM_AFTER As Long = &H1
Private Const LVCFMT_LEFT As Long = &H0
Private Const LVCFMT_RIGHT As Long = &H1
Private Const LVCFMT_CENTER As Long = &H2
Private Const LVCFMT_JUSTIFYMASK As Long = &H3
Private Const LVCFMT_FIXED_WIDTH As Long = &H100
Private Const LVCFMT_IMAGE As Long = &H800
Private Const LVCFMT_BITMAP_ON_RIGHT As Long = &H1000
Private Const LVCFMT_COL_HAS_IMAGES As Long = &H8000& ' Same as HDF_OWNERDRAW
Private Const LVCFMT_SPLITBUTTON As Long = &H1000000
Private Const LVTVIM_TILESIZE As Long = &H1
Private Const LVTVIM_COLUMNS As Long = &H2
Private Const LVTVIM_LABELMARGIN As Long = &H4
Private Const IIL_UNCHECKED As Long = 1
Private Const IIL_CHECKED As Long = 2
Private Const I_IMAGECALLBACK As Long = (-1)
Private Const I_COLUMNSCALLBACK As Long = (-1)
Private Const I_GROUPIDCALLBACK As Long = (-1)
Private Const I_GROUPIDNONE As Long = (-2)
Private Const MAX_PATH As Long = 260
Private Const NM_FIRST As Long = 0
Private Const NM_CLICK As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK As Long = (NM_FIRST - 3)
Private Const NM_RCLICK As Long = (NM_FIRST - 5)
Private Const NM_RDBLCLK As Long = (NM_FIRST - 6)
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Private Const LVS_ICON As Long = &H0
Private Const LVS_REPORT As Long = &H1
Private Const LVS_SMALLICON As Long = &H2
Private Const LVS_LIST As Long = &H3
Private Const LVS_TYPEMASK As Long = &H3
Private Const LVS_SINGLESEL As Long = &H4
Private Const LVS_SHOWSELALWAYS As Long = &H8
Private Const LVS_SORTASCENDING As Long = &H10
Private Const LVS_SORTDESCENDING As Long = &H20
Private Const LVS_SHAREIMAGELISTS As Long = &H40
Private Const LVS_NOLABELWRAP As Long = &H80
Private Const LVS_AUTOARRANGE As Long = &H100
Private Const LVS_EDITLABELS As Long = &H200
Private Const LVS_OWNERDATA As Long = &H1000
Private Const LVS_ALIGNTOP As Long = &H0
Private Const LVS_ALIGNLEFT As Long = &H800
Private Const LVS_OWNERDRAWFIXED As Long = &H400
Private Const LVS_NOCOLUMNHEADER As Long = &H4000
Private Const LVS_NOSORTHEADER As Long = &H8000&
Private Const LVS_EX_GRIDLINES As Long = &H1
Private Const LVS_EX_HEADERDRAGDROP As Long = &H10
Private Const LVS_EX_DOUBLEBUFFER As Long = &H10000
Private Const LVS_EX_SUBITEMIMAGES As Long = &H2
Private Const LVS_EX_FULLROWSELECT As Long = &H20
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_ONECLICKACTIVATE As Long = &H40
Private Const LVS_EX_INFOTIP As Long = &H400
Private Const LVS_EX_LABELTIP As Long = &H4000
Private Const LVS_EX_TRACKSELECT As Long = &H8
Private Const LVS_EX_UNDERLINEHOT As Long = &H800
Private Const LVS_EX_SNAPTOGRID As Long = &H80000
Private Const LV_VIEW_ICON As Long = &H0
Private Const LV_VIEW_DETAILS As Long = &H1
Private Const LV_VIEW_SMALLICON As Long = &H2
Private Const LV_VIEW_LIST As Long = &H3
Private Const LV_VIEW_TILE As Long = &H4
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private ListViewHandle As Long, ListViewHeaderHandle As Long, ListViewToolTipHandle As Long, ListViewHeaderToolTipHandle As Long
Private ListViewFontHandle As Long, ListViewBoldFontHandle As Long, ListViewUnderlineFontHandle As Long, ListViewBoldUnderlineFontHandle As Long
Private ListViewIMCHandle As Long
Private ListViewCharCodeCache As Long
Private ListViewIsClick As Boolean
Private ListViewMouseOver As Boolean
Private ListViewDesignMode As Boolean
Private ListViewFocusIndex As Long
Private ListViewLabelInEdit As Boolean
Private ListViewStartLabelEdit As Boolean
Private ListViewButtonDown As Integer
Private ListViewListItemsControl As Long
Private ListViewHotLightColor As Long
Private ListViewHotTrackItem As Long, ListViewHotTrackSubItem As Long
Private ListViewDragIndexBuffer As Long, ListViewDragIndex As Long
Private ListViewDragOffsetX As Long, ListViewDragOffsetY As Long
Private ListViewMemoryColumnWidth As Long
Private ListViewFilterEditHandle As Long, ListViewFilterEditIndex As Long
Private ListViewHeaderToolTipItem As Long
Private ListViewIconsObjectPointer As Long
Private ListViewSmallIconsObjectPointer As Long
Private ListViewColumnHeaderIconsObjectPointer As Long
Private ListViewGroupIconsObjectPointer As Long
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private DispIDHotMousePointer As Long
Private DispIDHeaderMousePointer As Long
Private DispIDIcons As Long, IconsArray() As String
Private DispIDSmallIcons As Long, SmallIconsArray() As String
Private DispIDColumnHeaderIcons As Long, ColumnHeaderIconsArray() As String
Private DispIDGroupIcons As Long, GroupIconsArray() As String
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropListItems As LvwListItems
Private PropColumnHeaders As LvwColumnHeaders
Private PropGroups As LvwGroups
Private PropWorkAreas As LvwWorkAreas
Private PropVisualStyles As Boolean
Private PropVisualTheme As LvwVisualThemeConstants
Private PropOLEDragMode As VBRUN.OLEDragConstants
Private PropOLEDragDropScroll As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropHotMousePointer As Integer, PropHotMouseIcon As IPictureDisp
Private PropHeaderMousePointer As Integer, PropHeaderMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropIconsName As String, PropIconsInit As Boolean
Private PropSmallIconsName As String, PropSmallIconsInit As Boolean
Private PropColumnHeaderIconsName As String, PropColumnHeaderIconsInit As Boolean
Private PropGroupIconsName As String, PropGroupIconsInit As Boolean
Private PropBorderStyle As CCBorderStyleConstants
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR
Private PropRedraw As Boolean
Private PropView As LvwViewConstants
Private PropArrange As LvwArrangeConstants
Private PropAllowColumnReorder As Boolean
Private PropAllowColumnCheckboxes As Boolean
Private PropMultiSelect As Boolean
Private PropFullRowSelect As Boolean
Private PropGridLines As Boolean
Private PropLabelEdit As LvwLabelEditConstants
Private PropLabelWrap As Boolean
Private PropSorted As Boolean
Private PropSortKey As Integer
Private PropSortOrder As LvwSortOrderConstants
Private PropSortType As LvwSortTypeConstants
Private PropCheckboxes As Boolean
Private PropHideSelection As Boolean
Private PropHideColumnHeaders As Boolean
Private PropShowInfoTips As Boolean
Private PropShowLabelTips As Boolean
Private PropShowColumnTips As Boolean
Private PropDoubleBuffer As Boolean
Private PropHoverSelection As Boolean
Private PropHoverSelectionTime As Long
Private PropHotTracking As Boolean
Private PropHighlightHot As Boolean
Private PropUnderlineHot As Boolean
Private PropInsertMarkColor As OLE_COLOR
Private PropTextBackground As CCBackStyleConstants
Private PropClickableColumnHeaders As Boolean
Private PropHighlightColumnHeaders As Boolean
Private PropTrackSizeColumnHeaders As Boolean
Private PropResizableColumnHeaders As Boolean
Private PropPicture As IPictureDisp
Private PropPictureAlignment As LvwPictureAlignmentConstants
Private PropPictureWatermark As Boolean
Private PropTileViewLines As Long
Private PropSnapToGrid As Boolean
Private PropGroupView As Boolean
Private PropGroupSubsetCount As Long
Private PropUseColumnChevron As Boolean
Private PropUseColumnFilterBar As Boolean
Private PropAutoSelectFirstItem As Boolean
Private PropIMEMode As CCIMEModeConstants
Private PropVirtualMode As Boolean
Private PropVirtualItemCount As Long
Private PropVirtualDisabledInfos As LvwVirtualPropertyConstants

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
            If ListViewLabelInEdit = False And ListViewFilterEditHandle = 0 Then
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
ElseIf DispID = DispIDHotMousePointer Then
    Call ComCtlsIPPBSetDisplayStringMousePointer(PropHotMousePointer, DisplayName)
    Handled = True
ElseIf DispID = DispIDHeaderMousePointer Then
    Call ComCtlsIPPBSetDisplayStringMousePointer(PropHeaderMousePointer, DisplayName)
    Handled = True
ElseIf DispID = DispIDIcons Then
    DisplayName = PropIconsName
    Handled = True
ElseIf DispID = DispIDSmallIcons Then
    DisplayName = PropSmallIconsName
    Handled = True
ElseIf DispID = DispIDColumnHeaderIcons Then
    DisplayName = PropColumnHeaderIconsName
    Handled = True
ElseIf DispID = DispIDGroupIcons Then
    DisplayName = PropGroupIconsName
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Or DispID = DispIDHotMousePointer Or DispID = DispIDHeaderMousePointer Then
    Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
    Handled = True
ElseIf DispID = DispIDIcons Or DispID = DispIDSmallIcons Or DispID = DispIDColumnHeaderIcons Or DispID = DispIDGroupIcons Then
    On Error GoTo CATCH_EXCEPTION
    Call ComCtlsIPPBSetPredefinedStringsImageList(StringsOut(), CookiesOut(), UserControl.ParentControls, IconsArray())
    SmallIconsArray() = IconsArray()
    ColumnHeaderIconsArray() = IconsArray()
    GroupIconsArray() = IconsArray()
    On Error GoTo 0
    Handled = True
End If
Exit Sub
CATCH_EXCEPTION:
Handled = False
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
If DispID = DispIDMousePointer Or DispID = DispIDHotMousePointer Or DispID = DispIDHeaderMousePointer Then
    Value = Cookie
    Handled = True
ElseIf DispID = DispIDIcons Then
    If Cookie < UBound(IconsArray()) Then Value = IconsArray(Cookie)
    Handled = True
ElseIf DispID = DispIDSmallIcons Then
    If Cookie < UBound(SmallIconsArray()) Then Value = SmallIconsArray(Cookie)
    Handled = True
ElseIf DispID = DispIDColumnHeaderIcons Then
    If Cookie < UBound(ColumnHeaderIconsArray()) Then Value = ColumnHeaderIconsArray(Cookie)
    Handled = True
ElseIf DispID = DispIDGroupIcons Then
    If Cookie < UBound(GroupIconsArray()) Then Value = GroupIconsArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_LISTVIEW_CLASSES)
Call SetVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ListViewHotLightColor = CLR_DEFAULT
ListViewHotTrackItem = -1
ListViewHotTrackSubItem = 0
ListViewHeaderToolTipItem = -1
ReDim IconsArray(0) As String
ReDim SmallIconsArray(0) As String
ReDim ColumnHeaderIconsArray(0) As String
ReDim GroupIconsArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDHotMousePointer = 0 Then DispIDHotMousePointer = GetDispID(Me, "HotMousePointer")
If DispIDHeaderMousePointer = 0 Then DispIDHeaderMousePointer = GetDispID(Me, "HeaderMousePointer")
If DispIDIcons = 0 Then DispIDIcons = GetDispID(Me, "Icons")
If DispIDSmallIcons = 0 Then DispIDSmallIcons = GetDispID(Me, "SmallIcons")
If DispIDColumnHeaderIcons = 0 Then DispIDColumnHeaderIcons = GetDispID(Me, "ColumnHeaderIcons")
If DispIDGroupIcons = 0 Then DispIDGroupIcons = GetDispID(Me, "GroupIcons")
On Error Resume Next
ListViewDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropVisualTheme = LvwVisualThemeStandard
PropOLEDragMode = vbOLEDragManual
PropOLEDragDropScroll = True
Me.OLEDropMode = vbOLEDropNone
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropHotMousePointer = 0: Set PropHotMouseIcon = Nothing
PropHeaderMousePointer = 0: Set PropHeaderMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropIconsName = "(None)"
PropSmallIconsName = "(None)"
PropColumnHeaderIconsName = "(None)"
PropGroupIconsName = "(None)"
PropBorderStyle = CCBorderStyleSunken
PropBackColor = vbWindowBackground
PropForeColor = vbWindowText
PropRedraw = True
PropView = LvwViewIcon
PropArrange = LvwArrangeNone
PropAllowColumnReorder = False
PropAllowColumnCheckboxes = False
PropMultiSelect = False
PropFullRowSelect = False
PropGridLines = False
PropLabelEdit = LvwLabelEditAutomatic
PropLabelWrap = True
PropSorted = False
PropSortKey = 0
PropSortOrder = LvwSortOrderAscending
PropSortType = LvwSortTypeBinary
PropCheckboxes = False
PropHideSelection = True
PropHideColumnHeaders = False
PropShowInfoTips = False
PropShowLabelTips = False
PropShowColumnTips = False
PropDoubleBuffer = True
PropHoverSelection = False
PropHoverSelectionTime = -1
PropHotTracking = False
PropHighlightHot = False
PropUnderlineHot = False
PropInsertMarkColor = vbBlack
PropTextBackground = CCBackStyleTransparent
PropClickableColumnHeaders = True
PropHighlightColumnHeaders = False
PropTrackSizeColumnHeaders = True
PropResizableColumnHeaders = True
Set PropPicture = Nothing
PropPictureAlignment = LvwPictureAlignmentTopLeft
PropPictureWatermark = False
PropTileViewLines = 1
PropSnapToGrid = False
PropGroupView = False
PropGroupSubsetCount = 0
PropUseColumnChevron = False
PropUseColumnFilterBar = False
PropAutoSelectFirstItem = True
PropIMEMode = CCIMEModeNoControl
PropVirtualMode = False
PropVirtualItemCount = 0
PropVirtualDisabledInfos = 0
Call CreateListView
If ListViewDesignMode = True Then
    Dim LVI As LVITEM, Buffer As String
    With LVI
    .Mask = LVIF_TEXT Or LVIF_INDENT
    .iItem = 0
    Buffer = Ambient.DisplayName
    .pszText = StrPtr(Buffer)
    .cchTextMax = Len(Buffer) + 1
    .iIndent = 0
    End With
    If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_INSERTITEM, 0, ByVal VarPtr(LVI)
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDHotMousePointer = 0 Then DispIDHotMousePointer = GetDispID(Me, "HotMousePointer")
If DispIDHeaderMousePointer = 0 Then DispIDHeaderMousePointer = GetDispID(Me, "HeaderMousePointer")
If DispIDIcons = 0 Then DispIDIcons = GetDispID(Me, "Icons")
If DispIDSmallIcons = 0 Then DispIDSmallIcons = GetDispID(Me, "SmallIcons")
If DispIDColumnHeaderIcons = 0 Then DispIDColumnHeaderIcons = GetDispID(Me, "ColumnHeaderIcons")
If DispIDGroupIcons = 0 Then DispIDGroupIcons = GetDispID(Me, "GroupIcons")
On Error Resume Next
ListViewDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
PropVisualTheme = .ReadProperty("VisualTheme", LvwVisualThemeStandard)
Me.Enabled = .ReadProperty("Enabled", True)
PropOLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
PropOLEDragDropScroll = .ReadProperty("OLEDragDropScroll", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropHotMousePointer = .ReadProperty("HotMousePointer", 0)
Set PropHotMouseIcon = .ReadProperty("HotMouseIcon", Nothing)
PropHeaderMousePointer = .ReadProperty("HeaderMousePointer", 0)
Set PropHeaderMouseIcon = .ReadProperty("HeaderMouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropIconsName = .ReadProperty("Icons", "(None)")
PropSmallIconsName = .ReadProperty("SmallIcons", "(None)")
PropColumnHeaderIconsName = .ReadProperty("ColumnHeaderIcons", "(None)")
PropGroupIconsName = .ReadProperty("GroupIcons", "(None)")
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleSunken)
PropBackColor = .ReadProperty("BackColor", vbWindowBackground)
PropForeColor = .ReadProperty("ForeColor", vbWindowText)
PropRedraw = .ReadProperty("Redraw", True)
PropView = .ReadProperty("View", LvwViewIcon)
PropArrange = .ReadProperty("Arrange", LvwArrangeNone)
PropAllowColumnReorder = .ReadProperty("AllowColumnReorder", False)
PropAllowColumnCheckboxes = .ReadProperty("AllowColumnCheckboxes", False)
PropMultiSelect = .ReadProperty("MultiSelect", False)
PropFullRowSelect = .ReadProperty("FullRowSelect", False)
PropGridLines = .ReadProperty("GridLines", False)
PropLabelEdit = .ReadProperty("LabelEdit", LvwLabelEditAutomatic)
PropLabelWrap = .ReadProperty("LabelWrap", True)
PropSorted = .ReadProperty("Sorted", False)
PropSortKey = .ReadProperty("SortKey", 0)
PropSortOrder = .ReadProperty("SortOrder", LvwSortOrderAscending)
PropSortType = .ReadProperty("SortType", LvwSortTypeBinary)
PropCheckboxes = .ReadProperty("Checkboxes", False)
PropHideSelection = .ReadProperty("HideSelection", True)
PropHideColumnHeaders = .ReadProperty("HideColumnHeaders", False)
PropShowInfoTips = .ReadProperty("ShowInfoTips", False)
PropShowLabelTips = .ReadProperty("ShowLabelTips", False)
PropShowColumnTips = .ReadProperty("ShowColumnTips", False)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
PropHoverSelection = .ReadProperty("HoverSelection", False)
PropHoverSelectionTime = .ReadProperty("HoverSelectionTime", -1)
PropHotTracking = .ReadProperty("HotTracking", False)
PropHighlightHot = .ReadProperty("HighlightHot", False)
PropUnderlineHot = .ReadProperty("UnderlineHot", False)
PropInsertMarkColor = .ReadProperty("InsertMarkColor", vbBlack)
PropTextBackground = .ReadProperty("TextBackground", CCBackStyleTransparent)
PropClickableColumnHeaders = .ReadProperty("ClickableColumnHeaders", True)
PropHighlightColumnHeaders = .ReadProperty("HighlightColumnHeaders", False)
PropTrackSizeColumnHeaders = .ReadProperty("TrackSizeColumnHeaders", True)
PropResizableColumnHeaders = .ReadProperty("ResizableColumnHeaders", True)
Set PropPicture = .ReadProperty("Picture", Nothing)
PropPictureAlignment = .ReadProperty("PictureAlignment", LvwPictureAlignmentTopLeft)
PropPictureWatermark = .ReadProperty("PictureWatermark", False)
PropTileViewLines = .ReadProperty("TileViewLines", 1)
PropSnapToGrid = .ReadProperty("SnapToGrid", False)
PropGroupView = .ReadProperty("GroupView", False)
PropGroupSubsetCount = .ReadProperty("GroupSubsetCount", 0)
PropUseColumnChevron = .ReadProperty("UseColumnChevron", False)
PropUseColumnFilterBar = .ReadProperty("UseColumnFilterBar", PropUseColumnFilterBar)
PropAutoSelectFirstItem = .ReadProperty("AutoSelectFirstItem", True)
PropIMEMode = .ReadProperty("IMEMode", CCIMEModeNoControl)
PropVirtualMode = .ReadProperty("VirtualMode", False)
PropVirtualItemCount = .ReadProperty("VirtualItemCount", 0)
PropVirtualDisabledInfos = .ReadProperty("VirtualDisabledInfos", 0)
End With
Call CreateListView
If ListViewDesignMode = False Then
    If Not PropIconsName = "(None)" Or Not PropSmallIconsName = "(None)" Or Not PropColumnHeaderIconsName = "(None)" Or Not PropGroupIconsName = "(None)" Then TimerImageList.Enabled = True
Else
    Dim LVI As LVITEM, Buffer As String
    With LVI
    .Mask = LVIF_TEXT Or LVIF_INDENT
    .iItem = 0
    Buffer = Ambient.DisplayName
    .pszText = StrPtr(Buffer)
    .cchTextMax = Len(Buffer) + 1
    .iIndent = 0
    End With
    If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_INSERTITEM, 0, ByVal VarPtr(LVI)
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "VisualTheme", PropVisualTheme, LvwVisualThemeStandard
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDragMode", PropOLEDragMode, vbOLEDragManual
.WriteProperty "OLEDragDropScroll", PropOLEDragDropScroll, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "HotMousePointer", PropHotMousePointer, 0
.WriteProperty "HotMouseIcon", PropHotMouseIcon, Nothing
.WriteProperty "HeaderMousePointer", PropHeaderMousePointer, 0
.WriteProperty "HeaderMouseIcon", PropHeaderMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "Icons", PropIconsName, "(None)"
.WriteProperty "SmallIcons", PropSmallIconsName, "(None)"
.WriteProperty "ColumnHeaderIcons", PropColumnHeaderIconsName, "(None)"
.WriteProperty "GroupIcons", PropGroupIconsName, "(None)"
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleSunken
.WriteProperty "BackColor", PropBackColor, vbWindowBackground
.WriteProperty "ForeColor", PropForeColor, vbWindowText
.WriteProperty "Redraw", PropRedraw, True
.WriteProperty "View", PropView, LvwViewIcon
.WriteProperty "Arrange", PropArrange, LvwArrangeNone
.WriteProperty "AllowColumnReorder", PropAllowColumnReorder, False
.WriteProperty "AllowColumnCheckboxes", PropAllowColumnCheckboxes, False
.WriteProperty "MultiSelect", PropMultiSelect, False
.WriteProperty "FullRowSelect", PropFullRowSelect, False
.WriteProperty "GridLines", PropGridLines, False
.WriteProperty "LabelEdit", PropLabelEdit, LvwLabelEditAutomatic
.WriteProperty "LabelWrap", PropLabelWrap, True
.WriteProperty "Sorted", PropSorted, False
.WriteProperty "SortKey", PropSortKey, 0
.WriteProperty "SortOrder", PropSortOrder, LvwSortOrderAscending
.WriteProperty "SortType", PropSortType, LvwSortTypeBinary
.WriteProperty "Checkboxes", PropCheckboxes, False
.WriteProperty "HideSelection", PropHideSelection, True
.WriteProperty "HideColumnHeaders", PropHideColumnHeaders, False
.WriteProperty "ShowInfoTips", PropShowInfoTips, False
.WriteProperty "ShowLabelTips", PropShowLabelTips, False
.WriteProperty "ShowColumnTips", PropShowColumnTips, False
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
.WriteProperty "HoverSelection", PropHoverSelection, False
.WriteProperty "HoverSelectionTime", PropHoverSelectionTime, -1
.WriteProperty "HotTracking", PropHotTracking, False
.WriteProperty "HighlightHot", PropHighlightHot, False
.WriteProperty "UnderlineHot", PropUnderlineHot, False
.WriteProperty "InsertMarkColor", PropInsertMarkColor, vbBlack
.WriteProperty "TextBackground", PropTextBackground, CCBackStyleTransparent
.WriteProperty "ClickableColumnHeaders", PropClickableColumnHeaders, True
.WriteProperty "HighlightColumnHeaders", PropHighlightColumnHeaders, False
.WriteProperty "TrackSizeColumnHeaders", PropTrackSizeColumnHeaders, True
.WriteProperty "ResizableColumnHeaders", PropResizableColumnHeaders, True
.WriteProperty "Picture", PropPicture, Nothing
.WriteProperty "PictureAlignment", PropPictureAlignment, LvwPictureAlignmentTopLeft
.WriteProperty "PictureWatermark", PropPictureWatermark, False
.WriteProperty "TileViewLines", PropTileViewLines, 1
.WriteProperty "SnapToGrid", PropSnapToGrid, False
.WriteProperty "GroupView", PropGroupView, False
.WriteProperty "GroupSubsetCount", PropGroupSubsetCount, 0
.WriteProperty "UseColumnChevron", PropUseColumnChevron, False
.WriteProperty "UseColumnFilterBar", PropUseColumnFilterBar, False
.WriteProperty "AutoSelectFirstItem", PropAutoSelectFirstItem, True
.WriteProperty "IMEMode", PropIMEMode, CCIMEModeNoControl
.WriteProperty "VirtualMode", PropVirtualMode, False
.WriteProperty "VirtualItemCount", PropVirtualItemCount, 0
.WriteProperty "VirtualDisabledInfos", PropVirtualDisabledInfos, 0
End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
ListViewDragIndex = 0
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
If ListViewHandle <> 0 Then
    If ListViewDragIndex > 0 And Not Effect = vbDropEffectNone Then
        Select Case PropView
            Case LvwViewIcon, LvwViewSmallIcon, LvwViewTile
                Select Case PropArrange
                    Case LvwArrangeNone, LvwArrangeLeft, LvwArrangeTop
                    Case Else
                        Effect = vbDropEffectNone
                End Select
        End Select
    End If
    If State = vbOver And Not Effect = vbDropEffectNone Then
        If PropOLEDragDropScroll = True Then
            Dim RC As RECT
            GetWindowRect ListViewHandle, RC
            Dim dwStyle As Long
            dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
            If (dwStyle And WS_HSCROLL) = WS_HSCROLL Then
                If Abs(X) < (16 * PixelsPerDIP_X()) Then
                    SendMessage ListViewHandle, WM_HSCROLL, SB_LINELEFT, ByVal 0&
                ElseIf Abs(X - (RC.Right - RC.Left)) < (16 * PixelsPerDIP_X()) Then
                    SendMessage ListViewHandle, WM_HSCROLL, SB_LINERIGHT, ByVal 0&
                End If
            End If
            If (dwStyle And WS_VSCROLL) = WS_VSCROLL Then
                If Abs(Y) < (16 * PixelsPerDIP_Y()) Then
                    SendMessage ListViewHandle, WM_VSCROLL, SB_LINEUP, ByVal 0&
                ElseIf Abs(Y - (RC.Bottom - RC.Top)) < (16 * PixelsPerDIP_Y()) Then
                    SendMessage ListViewHandle, WM_VSCROLL, SB_LINEDOWN, ByVal 0&
                End If
            End If
        End If
    End If
    If ListViewDragIndex > 0 And Not Effect = vbDropEffectNone Then
        Select Case PropView
            Case LvwViewIcon, LvwViewSmallIcon, LvwViewTile
                Select Case PropArrange
                    Case LvwArrangeNone, LvwArrangeLeft, LvwArrangeTop
                        Dim ViewRect As RECT, P As POINTAPI
                        SendMessage ListViewHandle, LVM_GETVIEWRECT, 0, ByVal VarPtr(ViewRect)
                        If (CDbl(X) + (CDbl(ListViewDragOffsetX) - CDbl(ViewRect.Left))) <= MAXINT_4 Then P.X = CLng(CDbl(X) + (CDbl(ListViewDragOffsetX) - CDbl(ViewRect.Left))) Else P.X = MAXINT_4
                        If (CDbl(Y) + (CDbl(ListViewDragOffsetY) - CDbl(ViewRect.Top))) <= MAXINT_4 Then P.Y = CLng(CDbl(Y) + (CDbl(ListViewDragOffsetY) - CDbl(ViewRect.Top))) Else P.Y = MAXINT_4
                        SendMessage ListViewHandle, LVM_SETITEMPOSITION32, ListViewDragIndex - 1, ByVal VarPtr(P)
                End Select
        End Select
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
If ListViewDragIndex > 0 Then
    If ListViewHandle <> 0 Then
        Dim P(0 To 1) As POINTAPI, RC As RECT
        GetCursorPos P(0)
        ScreenToClient ListViewHandle, P(0)
        SendMessage ListViewHandle, LVM_GETITEMPOSITION, ListViewDragIndex - 1, ByVal VarPtr(P(1))
        SendMessage ListViewHandle, LVM_GETVIEWRECT, 0, ByVal VarPtr(RC)
        If ((CDbl(P(1).X) - CDbl(P(0).X)) + CDbl(RC.Left)) <= MAXINT_4 Then ListViewDragOffsetX = CLng((CDbl(P(1).X) - CDbl(P(0).X)) + CDbl(RC.Left)) Else ListViewDragOffsetX = MAXINT_4
        If ((CDbl(P(1).Y) - CDbl(P(0).Y)) + CDbl(RC.Top)) <= MAXINT_4 Then ListViewDragOffsetY = CLng((CDbl(P(1).Y) - CDbl(P(0).Y)) + CDbl(RC.Top)) Else ListViewDragOffsetY = MAXINT_4
    End If
    If PropOLEDragMode = vbOLEDragAutomatic Then
        Dim Text As String
        Text = Me.FListItemText(ListViewDragIndex, 0)
        Data.SetData StrToVar(Text & vbNullChar), CF_UNICODETEXT
        Data.SetData Text, vbCFText
        AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    End If
End If
RaiseEvent OLEStartDrag(Data, AllowedEffects)
If AllowedEffects = vbDropEffectNone Then ListViewDragIndex = 0
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
If ListViewDragIndex > 0 Then Exit Sub
If ListViewDragIndexBuffer > 0 Then ListViewDragIndex = ListViewDragIndexBuffer
UserControl.OLEDrag
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If ListViewDesignMode = True And PropertyName = "DisplayName" Then
    If ListViewHandle <> 0 Then
        If SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&) > 0 Then Me.FListItemText(1, 0) = Ambient.DisplayName
    End If
End If
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If ListViewHandle <> 0 Then MoveWindow ListViewHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyListView
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropIconsInit = False Then
    If Not PropIconsName = "(None)" Then Me.Icons = PropIconsName
    PropIconsInit = True
End If
If PropSmallIconsInit = False Then
    If Not PropSmallIconsName = "(None)" Then Me.SmallIcons = PropSmallIconsName
    PropSmallIconsInit = True
End If
If PropColumnHeaderIconsInit = False Then
    If Not PropColumnHeaderIconsName = "(None)" Then Me.ColumnHeaderIcons = PropColumnHeaderIconsName
    PropColumnHeaderIconsInit = True
End If
If PropGroupIconsInit = False Then
    If Not PropGroupIconsName = "(None)" Then Me.GroupIcons = PropGroupIconsName
    PropGroupIconsInit = True
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
hWnd = ListViewHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get hWndHeader() As Long
Attribute hWndHeader.VB_Description = "Returns a handle to a control."
If ListViewHandle <> 0 Then hWndHeader = SendMessage(ListViewHandle, LVM_GETHEADER, 0, ByVal 0&)
End Property

Public Property Get hWndLabelEdit() As Long
Attribute hWndLabelEdit.VB_Description = "Returns a handle to a control."
If ListViewHandle <> 0 Then hWndLabelEdit = SendMessage(ListViewHandle, LVM_GETEDITCONTROL, 0, ByVal 0&)
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
Dim OldFontHandle As Long, OldBoldFontHandle As Long, OldUnderlineFontHandle As Long, OldBoldUnderlineFontHandle As Long
Dim TempFont As StdFont
Set PropFont = NewFont
OldFontHandle = ListViewFontHandle
OldBoldFontHandle = ListViewBoldFontHandle
OldUnderlineFontHandle = ListViewUnderlineFontHandle
OldBoldUnderlineFontHandle = ListViewBoldUnderlineFontHandle
ListViewFontHandle = CreateGDIFontFromOLEFont(PropFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Bold = True
ListViewBoldFontHandle = CreateGDIFontFromOLEFont(TempFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Underline = True
ListViewUnderlineFontHandle = CreateGDIFontFromOLEFont(TempFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Bold = True
TempFont.Underline = True
ListViewBoldUnderlineFontHandle = CreateGDIFontFromOLEFont(TempFont)
If ListViewHandle <> 0 Then SendMessage ListViewHandle, WM_SETFONT, ListViewFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If OldBoldFontHandle <> 0 Then DeleteObject OldBoldFontHandle
If OldUnderlineFontHandle <> 0 Then DeleteObject OldUnderlineFontHandle
If OldBoldUnderlineFontHandle <> 0 Then DeleteObject OldBoldUnderlineFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long, OldBoldFontHandle As Long, OldUnderlineFontHandle As Long, OldBoldUnderlineFontHandle As Long
Dim TempFont As StdFont
OldFontHandle = ListViewFontHandle
OldBoldFontHandle = ListViewBoldFontHandle
OldUnderlineFontHandle = ListViewUnderlineFontHandle
OldBoldUnderlineFontHandle = ListViewBoldUnderlineFontHandle
ListViewFontHandle = CreateGDIFontFromOLEFont(PropFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Bold = True
ListViewBoldFontHandle = CreateGDIFontFromOLEFont(TempFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Underline = True
ListViewUnderlineFontHandle = CreateGDIFontFromOLEFont(TempFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Bold = True
TempFont.Underline = True
ListViewBoldUnderlineFontHandle = CreateGDIFontFromOLEFont(TempFont)
If ListViewHandle <> 0 Then SendMessage ListViewHandle, WM_SETFONT, ListViewFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If OldBoldFontHandle <> 0 Then DeleteObject OldBoldFontHandle
If OldUnderlineFontHandle <> 0 Then DeleteObject OldUnderlineFontHandle
If OldBoldUnderlineFontHandle <> 0 Then DeleteObject OldBoldUnderlineFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If ListViewHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        If PropVisualTheme = LvwVisualThemeExplorer Then
            SetWindowTheme ListViewHandle, StrPtr("Explorer"), 0
        Else
            ActivateVisualStyles ListViewHandle
        End If
    Else
        RemoveVisualStyles ListViewHandle
    End If
    Call SetVisualStylesHeader
    Call SetVisualStylesToolTip
    Call SetVisualStylesHeaderToolTip
    Me.Refresh
    If ComCtlsSupportLevel() >= 2 Then
        If Not PropPicture Is Nothing Then
            If PropPictureAlignment = LvwPictureAlignmentTile Then Set Me.Picture = PropPicture
        End If
    End If
End If
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get VisualTheme() As LvwVisualThemeConstants
Attribute VisualTheme.VB_Description = "Returns/sets the visual theme. Requires comctl32.dll version 6.0 or higher."
VisualTheme = PropVisualTheme
End Property

Public Property Let VisualTheme(ByVal Value As LvwVisualThemeConstants)
Select Case Value
    Case LvwVisualThemeStandard, LvwVisualThemeExplorer
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
If ListViewHandle <> 0 Then EnableWindow ListViewHandle, IIf(Value = True, 1, 0)
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
If ListViewDesignMode = False Then Call RefreshMousePointer
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
        If ListViewDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ListViewDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get HotMousePointer() As Integer
Attribute HotMousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over an item while hot tracking is enabled."
HotMousePointer = PropHotMousePointer
End Property

Public Property Let HotMousePointer(ByVal Value As Integer)
Select Case Value
    Case 0 To 16, 99
        PropHotMousePointer = Value
        If ListViewHandle <> 0 Then
            If MousePointerID(PropHotMousePointer) <> 0 Then
                SendMessage ListViewHandle, LVM_SETHOTCURSOR, 0, ByVal LoadCursor(0, MousePointerID(PropHotMousePointer))
            ElseIf PropHotMousePointer = 99 And Not PropHotMouseIcon Is Nothing Then
                SendMessage ListViewHandle, LVM_SETHOTCURSOR, 0, ByVal PropHotMouseIcon.Handle
            Else
                SendMessage ListViewHandle, LVM_SETHOTCURSOR, 0, ByVal 0&
            End If
        End If
    Case Else
        Err.Raise 380
End Select
If ListViewDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "HotMousePointer"
End Property

Public Property Get HotMouseIcon() As IPictureDisp
Attribute HotMouseIcon.VB_Description = "Returns/sets a custom hot mouse icon."
Set HotMouseIcon = PropHotMouseIcon
End Property

Public Property Let HotMouseIcon(ByVal Value As IPictureDisp)
Set Me.HotMouseIcon = Value
End Property

Public Property Set HotMouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropHotMouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
        Set PropHotMouseIcon = Value
    Else
        If ListViewDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ListViewHandle <> 0 Then
    If MousePointerID(PropHotMousePointer) <> 0 Then
        SendMessage ListViewHandle, LVM_SETHOTCURSOR, 0, ByVal LoadCursor(0, MousePointerID(PropHotMousePointer))
    ElseIf PropHotMousePointer = 99 And Not PropHotMouseIcon Is Nothing Then
        SendMessage ListViewHandle, LVM_SETHOTCURSOR, 0, ByVal PropHotMouseIcon.Handle
    Else
        SendMessage ListViewHandle, LVM_SETHOTCURSOR, 0, ByVal 0&
    End If
End If
If ListViewDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "HotMouseIcon"
End Property

Public Property Get HeaderMousePointer() As Integer
Attribute HeaderMousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over the column headers."
HeaderMousePointer = PropHeaderMousePointer
End Property

Public Property Let HeaderMousePointer(ByVal Value As Integer)
Select Case Value
    Case 0 To 16, 99
        PropHeaderMousePointer = Value
    Case Else
        Err.Raise 380
End Select
If ListViewDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "HeaderMousePointer"
End Property

Public Property Get HeaderMouseIcon() As IPictureDisp
Attribute HeaderMouseIcon.VB_Description = "Returns/sets a custom header mouse icon."
Set HeaderMouseIcon = PropHeaderMouseIcon
End Property

Public Property Let HeaderMouseIcon(ByVal Value As IPictureDisp)
Set Me.HeaderMouseIcon = Value
End Property

Public Property Set HeaderMouseIcon(ByVal Value As IPictureDisp)
If Value Is Nothing Then
    Set PropHeaderMouseIcon = Nothing
Else
    If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
        Set PropHeaderMouseIcon = Value
    Else
        If ListViewDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ListViewDesignMode = False Then Call RefreshMousePointer
UserControl.PropertyChanged "HeaderMouseIcon"
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
If ListViewDesignMode = False Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
    dwMask = 0
End If
If PropRightToLeft = True Then
    If PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = WS_EX_RTLREADING
End If
If ListViewHandle <> 0 Then Call ComCtlsSetRightToLeft(ListViewHandle, dwMask)
If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
If ListViewHeaderHandle <> 0 Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = 0
    Call ComCtlsSetRightToLeft(ListViewHeaderHandle, dwMask)
    If Me.ColumnHeaders.Count > 0 Then
        Dim i As Long
        For i = 1 To Me.ColumnHeaders.Count
            Call SetColumnRTLReading(i, CBool(PropRightToLeft = True And PropRightToLeftLayout = False))
        Next i
    End If
End If
If ListViewToolTipHandle <> 0 Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = 0
    Call ComCtlsSetRightToLeft(ListViewToolTipHandle, dwMask)
End If
If ListViewHeaderToolTipHandle <> 0 Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = 0
    Call ComCtlsSetRightToLeft(ListViewHeaderToolTipHandle, dwMask)
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

Public Property Get Icons() As Variant
Attribute Icons.VB_Description = "Returns/sets the image list control to be used for the icons."
If ListViewDesignMode = False Then
    If PropIconsInit = False And ListViewIconsObjectPointer = 0 Then
        If Not PropIconsName = "(None)" Then Me.Icons = PropIconsName
        PropIconsInit = True
    End If
    Set Icons = PropIconsControl
Else
    Icons = PropIconsName
End If
End Property

Public Property Set Icons(ByVal Value As Variant)
Me.Icons = Value
End Property

Public Property Let Icons(ByVal Value As Variant)
If ListViewDesignMode = False Then
    If ListViewHandle <> 0 Then
        Dim Success As Boolean, Handle As Long
        On Error Resume Next
        If IsObject(Value) Then
            If TypeName(Value) = "ImageList" Then
                Handle = Value.hImageList
                Success = CBool(Err.Number = 0 And Handle <> 0)
            End If
            If Success = True Then
                SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal Handle
                ListViewIconsObjectPointer = ObjPtr(Value)
                PropIconsName = ProperControlName(Value)
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
                            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal Handle
                            ListViewIconsObjectPointer = ObjPtr(ControlEnum)
                            PropIconsName = Value
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal 0&
            ListViewIconsObjectPointer = 0
            PropIconsName = "(None)"
        ElseIf Handle = 0 Then
            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_NORMAL, ByVal 0&
        Else
            SendMessage ListViewHandle, LVM_ARRANGE, LVA_DEFAULT, ByVal 0&
        End If
    End If
Else
    PropIconsName = Value
End If
UserControl.PropertyChanged "Icons"
End Property

Public Property Get SmallIcons() As Variant
Attribute SmallIcons.VB_Description = "Returns/sets the image list control to be used for the small icons."
If ListViewDesignMode = False Then
    If PropSmallIconsInit = False And ListViewSmallIconsObjectPointer = 0 Then
        If Not PropSmallIconsName = "(None)" Then Me.SmallIcons = PropSmallIconsName
        PropSmallIconsInit = True
    End If
    Set SmallIcons = PropSmallIconsControl
Else
    SmallIcons = PropSmallIconsName
End If
End Property

Public Property Set SmallIcons(ByVal Value As Variant)
Me.SmallIcons = Value
End Property

Public Property Let SmallIcons(ByVal Value As Variant)
If ListViewDesignMode = False Then
    If ListViewHandle <> 0 Then
        Dim Success As Boolean, Handle As Long, Size As SIZEAPI
        If PropView = LvwViewList Then
            ListViewMemoryColumnWidth = SendMessage(ListViewHandle, LVM_GETCOLUMNWIDTH, 0, ByVal 0&)
            Handle = SendMessage(ListViewHandle, LVM_GETIMAGELIST, LVSIL_SMALL, ByVal 0&)
            If Handle <> 0 Then
                ImageList_GetIconSize Handle, Size.CX, Size.CY
                ListViewMemoryColumnWidth = ListViewMemoryColumnWidth - Size.CX
                Handle = 0
            End If
        End If
        On Error Resume Next
        If IsObject(Value) Then
            If TypeName(Value) = "ImageList" Then
                Handle = Value.hImageList
                Success = CBool(Err.Number = 0 And Handle <> 0)
            End If
            If Success = True Then
                SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal Handle
                ListViewSmallIconsObjectPointer = ObjPtr(Value)
                PropSmallIconsName = ProperControlName(Value)
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
                            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal Handle
                            ListViewSmallIconsObjectPointer = ObjPtr(ControlEnum)
                            PropSmallIconsName = Value
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal 0&
            ListViewSmallIconsObjectPointer = 0
            PropSmallIconsName = "(None)"
        ElseIf Handle = 0 Then
            SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal 0&
        Else
            SendMessage ListViewHandle, LVM_ARRANGE, LVA_DEFAULT, ByVal 0&
            If PropView = LvwViewList Then
                ImageList_GetIconSize Handle, Size.CX, Size.CY
                ListViewMemoryColumnWidth = ListViewMemoryColumnWidth + Size.CX
                If ListViewMemoryColumnWidth > 0 Then SendMessage ListViewHandle, LVM_SETCOLUMNWIDTH, 0, ByVal ListViewMemoryColumnWidth
            End If
        End If
        ' The image list for the column icons need to be reset, because
        ' LVM_SETIMAGELIST with LVSIL_SMALL overrides the image list for the column icons.
        If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
        If ListViewHeaderHandle <> 0 Then
            If Not PropColumnHeaderIconsControl Is Nothing Then
                Dim ImageListHandle As Long
                ImageListHandle = PropColumnHeaderIconsControl.hImageList
                SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal ImageListHandle
                RedrawWindow ListViewHeaderHandle, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
            Else
                SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal 0&
            End If
        End If
    End If
Else
    PropSmallIconsName = Value
End If
UserControl.PropertyChanged "SmallIcons"
End Property

Public Property Get ColumnHeaderIcons() As Variant
Attribute ColumnHeaderIcons.VB_Description = "Returns/sets the image list control to be used for the column header icons."
If ListViewDesignMode = False Then
    If PropColumnHeaderIconsInit = False And ListViewColumnHeaderIconsObjectPointer = 0 Then
        If Not PropColumnHeaderIconsName = "(None)" Then Me.ColumnHeaderIcons = PropColumnHeaderIconsName
        PropColumnHeaderIconsInit = True
    End If
    Set ColumnHeaderIcons = PropColumnHeaderIconsControl
Else
    ColumnHeaderIcons = PropColumnHeaderIconsName
End If
End Property

Public Property Set ColumnHeaderIcons(ByVal Value As Variant)
Me.ColumnHeaderIcons = Value
End Property

Public Property Let ColumnHeaderIcons(ByVal Value As Variant)
If ListViewDesignMode = False Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHandle <> 0 And ListViewHeaderHandle <> 0 Then
        Dim Success As Boolean, Handle As Long
        On Error Resume Next
        If IsObject(Value) Then
            If TypeName(Value) = "ImageList" Then
                Handle = Value.hImageList
                Success = CBool(Err.Number = 0 And Handle <> 0)
            End If
            If Success = True Then
                SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal Handle
                ListViewColumnHeaderIconsObjectPointer = ObjPtr(Value)
                PropColumnHeaderIconsName = ProperControlName(Value)
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
                            SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal Handle
                            ListViewColumnHeaderIconsObjectPointer = ObjPtr(ControlEnum)
                            PropColumnHeaderIconsName = Value
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal 0&
            ListViewColumnHeaderIconsObjectPointer = 0
            PropColumnHeaderIconsName = "(None)"
        ElseIf Handle = 0 Then
            SendMessage ListViewHeaderHandle, HDM_SETIMAGELIST, HDSIL_NORMAL, ByVal 0&
        End If
        If Me.ColumnHeaders.Count > 0 Then
            Dim i As Long, Icon As Long
            For i = 1 To Me.ColumnHeaders.Count
                Icon = Me.FColumnHeaderIcon(i)
                If Icon > 0 Then Me.FColumnHeaderIcon(i) = Icon
            Next i
        End If
    End If
Else
    PropColumnHeaderIconsName = Value
End If
UserControl.PropertyChanged "ColumnHeaderIcons"
End Property

Public Property Get GroupIcons() As Variant
Attribute GroupIcons.VB_Description = "Returns/sets the image list control to be used for the group header icons. Requires comctl32.dll version 6.1 or higher."
If ListViewDesignMode = False Then
    If PropGroupIconsInit = False And ListViewGroupIconsObjectPointer = 0 Then
        If Not PropGroupIconsName = "(None)" Then Me.GroupIcons = PropGroupIconsName
        PropGroupIconsInit = True
    End If
    Set GroupIcons = PropGroupIconsControl
Else
    GroupIcons = PropGroupIconsName
End If
End Property

Public Property Set GroupIcons(ByVal Value As Variant)
Me.GroupIcons = Value
End Property

Public Property Let GroupIcons(ByVal Value As Variant)
If ListViewDesignMode = False Then
    If ListViewHandle <> 0 Then
        Dim Success As Boolean, Handle As Long
        On Error Resume Next
        If IsObject(Value) Then
            If TypeName(Value) = "ImageList" Then
                Handle = Value.hImageList
                Success = CBool(Err.Number = 0 And Handle <> 0)
            End If
            If Success = True Then
                If ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_GROUPHEADER, ByVal Handle
                ListViewGroupIconsObjectPointer = ObjPtr(Value)
                PropGroupIconsName = ProperControlName(Value)
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
                            If ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_GROUPHEADER, ByVal Handle
                            ListViewGroupIconsObjectPointer = ObjPtr(ControlEnum)
                            PropGroupIconsName = Value
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            If ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_GROUPHEADER, ByVal 0&
            ListViewGroupIconsObjectPointer = 0
            PropGroupIconsName = "(None)"
        ElseIf Handle = 0 Then
            If ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_SETIMAGELIST, LVSIL_GROUPHEADER, ByVal 0&
        End If
    End If
Else
    PropGroupIconsName = Value
End If
UserControl.PropertyChanged "GroupIcons"
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
If ListViewHandle <> 0 Then Call ComCtlsChangeBorderStyle(ListViewHandle, PropBorderStyle)
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
If Value = CLR_NONE Then Err.Raise 380
PropBackColor = Value
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, LVM_SETBKCOLOR, 0, ByVal WinColor(PropBackColor)
    If PropTextBackground = CCBackStyleOpaque Then SendMessage ListViewHandle, LVM_SETTEXTBKCOLOR, 0, ByVal WinColor(PropBackColor)
    Me.Refresh
    If Not PropPicture Is Nothing Then
        If PropPicture.Type = vbPicTypeIcon Then Set Me.Picture = PropPicture
    End If
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
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, LVM_SETTEXTCOLOR, 0, ByVal WinColor(PropForeColor)
    Me.Refresh
End If
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Returns/sets a value that determines whether or not the list view redraws when changing the list items. You can speed up the creation of large lists by disabling this property before adding the list items."
Redraw = PropRedraw
End Property

Public Property Let Redraw(ByVal Value As Boolean)
PropRedraw = Value
If ListViewHandle <> 0 And ListViewDesignMode = False Then
    SendMessage ListViewHandle, WM_SETREDRAW, IIf(PropRedraw = True, 1, 0), ByVal 0&
    If PropRedraw = True Then Me.Refresh
End If
UserControl.PropertyChanged "Redraw"
End Property

Public Property Get View() As LvwViewConstants
Attribute View.VB_Description = "Returns/sets the current view."
View = PropView
End Property

Public Property Let View(ByVal Value As LvwViewConstants)
Select Case Value
    Case LvwViewIcon, LvwViewSmallIcon, LvwViewList, LvwViewReport, LvwViewTile
        If PropVirtualMode = True And Value = LvwViewTile Then
            If ListViewDesignMode = True Then
                MsgBox "View must not be 4 - Tile when VirtualMode is True", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise Number:=383, Description:="View must not be 4 - Tile when VirtualMode is True"
            End If
        End If
        PropView = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 And ListViewDesignMode = False Then
    If ComCtlsSupportLevel() >= 1 Then
        Dim NewView As Long
        Select Case PropView
            Case LvwViewIcon
                NewView = LV_VIEW_ICON
            Case LvwViewSmallIcon
                NewView = LV_VIEW_SMALLICON
            Case LvwViewList
                NewView = LV_VIEW_LIST
            Case LvwViewReport
                NewView = LV_VIEW_DETAILS
            Case LvwViewTile
                NewView = LV_VIEW_TILE
        End Select
        SendMessage ListViewHandle, LVM_SETVIEW, NewView, ByVal 0&
    Else
        If PropView = LvwViewTile Then PropView = LvwViewIcon
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
        dwStyle = dwStyle And Not LVS_TYPEMASK
        Select Case PropView
            Case LvwViewIcon
                dwStyle = dwStyle Or LVS_ICON
            Case LvwViewSmallIcon
                dwStyle = dwStyle Or LVS_SMALLICON
            Case LvwViewList
                dwStyle = dwStyle Or LVS_LIST
            Case LvwViewReport
                dwStyle = dwStyle Or LVS_REPORT
        End Select
        SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
    End If
    If PropView = LvwViewList Then
        If ListViewMemoryColumnWidth > 0 Then SendMessage ListViewHandle, LVM_SETCOLUMNWIDTH, 0, ByVal ListViewMemoryColumnWidth
    ElseIf PropView = LvwViewReport Then
        Call CheckHeaderControl
    End If
    If ComCtlsSupportLevel() >= 2 Then
        If Not PropPicture Is Nothing Then
            If PropPictureAlignment = LvwPictureAlignmentTile Then Set Me.Picture = PropPicture
        End If
    End If
End If
UserControl.PropertyChanged "View"
End Property

Public Property Get Arrange() As LvwArrangeConstants
Attribute Arrange.VB_Description = "Returns/sets a value indicating how the icons in a 'icon', 'small icon' or 'tile' view are arranged."
Arrange = PropArrange
End Property

Public Property Let Arrange(ByVal Value As LvwArrangeConstants)
Select Case Value
    Case LvwArrangeNone, LvwArrangeAutoLeft, LvwArrangeAutoTop, LvwArrangeLeft, LvwArrangeTop
        If PropVirtualMode = True And Value <> LvwArrangeNone Then
            If ListViewDesignMode = True Then
                MsgBox "Arrange must be 0 - None when VirtualMode is True", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise Number:=383, Description:="Arrange must be 0 - None when VirtualMode is True"
            End If
        End If
        PropArrange = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 And ListViewDesignMode = False Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If (dwStyle And LVS_AUTOARRANGE) = LVS_AUTOARRANGE Then dwStyle = dwStyle And Not LVS_AUTOARRANGE
    If (dwStyle And LVS_ALIGNLEFT) = LVS_ALIGNLEFT Then dwStyle = dwStyle And Not LVS_ALIGNLEFT
    If (dwStyle And LVS_ALIGNTOP) = LVS_ALIGNTOP Then dwStyle = dwStyle And Not LVS_ALIGNTOP
    If PropVirtualMode = False Then
        Select Case PropArrange
            Case LvwArrangeAutoLeft
                dwStyle = dwStyle Or LVS_AUTOARRANGE Or LVS_ALIGNLEFT
            Case LvwArrangeAutoTop
                dwStyle = dwStyle Or LVS_AUTOARRANGE Or LVS_ALIGNTOP
            Case LvwArrangeLeft
                dwStyle = dwStyle Or LVS_ALIGNLEFT
            Case LvwArrangeTop
                dwStyle = dwStyle Or LVS_ALIGNTOP
        End Select
    Else
        ' According to MSDN:
        ' All virtual list view controls default to the LVS_AUTOARRANGE style.
        dwStyle = dwStyle Or LVS_AUTOARRANGE
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "Arrange"
End Property

Public Property Get AllowColumnReorder() As Boolean
Attribute AllowColumnReorder.VB_Description = "Returns/sets a value that determines whether or not a user can reorder column headers in 'report' view."
AllowColumnReorder = PropAllowColumnReorder
End Property

Public Property Let AllowColumnReorder(ByVal Value As Boolean)
PropAllowColumnReorder = Value
If ListViewHandle <> 0 Then
    If PropAllowColumnReorder = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HEADERDRAGDROP, ByVal LVS_EX_HEADERDRAGDROP
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_HEADERDRAGDROP, ByVal 0&
    End If
End If
UserControl.PropertyChanged "AllowColumnReorder"
End Property

Public Property Get AllowColumnCheckboxes() As Boolean
Attribute AllowColumnCheckboxes.VB_Description = "Returns/sets a value that determines whether or not the column headers in 'report' view are allowed to place checkboxes. Requires comctl32.dll version 6.1 or higher."
AllowColumnCheckboxes = PropAllowColumnCheckboxes
End Property

Public Property Let AllowColumnCheckboxes(ByVal Value As Boolean)
PropAllowColumnCheckboxes = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropAllowColumnCheckboxes = CBool((dwStyle And HDS_CHECKBOXES) = HDS_CHECKBOXES) Then
            If PropAllowColumnCheckboxes = True Then
                If Not (dwStyle And HDS_CHECKBOXES) = HDS_CHECKBOXES Then dwStyle = dwStyle Or HDS_CHECKBOXES
            Else
                If (dwStyle And HDS_CHECKBOXES) = HDS_CHECKBOXES Then dwStyle = dwStyle And Not HDS_CHECKBOXES
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "AllowColumnCheckboxes"
End Property

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the list view and how the multiple selections can be made."
MultiSelect = PropMultiSelect
End Property

Public Property Let MultiSelect(ByVal Value As Boolean)
PropMultiSelect = Value
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If PropMultiSelect = True Then
        If (dwStyle And LVS_SINGLESEL) = LVS_SINGLESEL Then dwStyle = dwStyle And Not LVS_SINGLESEL
    Else
        If Not (dwStyle And LVS_SINGLESEL) = LVS_SINGLESEL Then dwStyle = dwStyle Or LVS_SINGLESEL
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
    If PropMultiSelect = False Then
        Dim ListItem As LvwListItem
        Set ListItem = Me.SelectedItem
        If Not ListItem Is Nothing Then ListItem.Selected = True
    End If
End If
UserControl.PropertyChanged "MultiSelect"
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets whether selecting a list item highlights the entire row in 'report' view."
FullRowSelect = PropFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal Value As Boolean)
PropFullRowSelect = Value
If ListViewHandle <> 0 Then
    If PropFullRowSelect = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal LVS_EX_FULLROWSELECT
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal 0&
    End If
End If
UserControl.PropertyChanged "FullRowSelect"
End Property

Public Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "Returns/sets whether grid lines appear between rows and columns in 'report' view."
GridLines = PropGridLines
End Property

Public Property Let GridLines(ByVal Value As Boolean)
PropGridLines = Value
If ListViewHandle <> 0 Then
    If PropGridLines = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, ByVal LVS_EX_GRIDLINES
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, ByVal 0&
    End If
End If
UserControl.PropertyChanged "GridLines"
End Property

Public Property Get LabelEdit() As LvwLabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a list item."
LabelEdit = PropLabelEdit
End Property

Public Property Let LabelEdit(ByVal Value As LvwLabelEditConstants)
Select Case Value
    Case LvwLabelEditAutomatic, LvwLabelEditManual, LvwLabelEditDisabled
        PropLabelEdit = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    Select Case PropLabelEdit
        Case LvwLabelEditAutomatic, LvwLabelEditManual
            If Not (dwStyle And LVS_EDITLABELS) = LVS_EDITLABELS Then dwStyle = dwStyle Or LVS_EDITLABELS
        Case LvwLabelEditDisabled
            If (dwStyle And LVS_EDITLABELS) = LVS_EDITLABELS Then dwStyle = dwStyle And Not LVS_EDITLABELS
    End Select
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "LabelEdit"
End Property

Public Property Get LabelWrap() As Boolean
Attribute LabelWrap.VB_Description = "Returns/sets a value that determines if labels are wrapped when the list view is in icon view."
LabelWrap = PropLabelWrap
End Property

Public Property Let LabelWrap(ByVal Value As Boolean)
PropLabelWrap = Value
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If PropLabelWrap = True Then
        If (dwStyle And LVS_NOLABELWRAP) = LVS_NOLABELWRAP Then dwStyle = dwStyle And Not LVS_NOLABELWRAP
    Else
        If Not (dwStyle And LVS_NOLABELWRAP) = LVS_NOLABELWRAP Then dwStyle = dwStyle Or LVS_NOLABELWRAP
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "LabelWrap"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Returns/sets a value indicating if the list items are automatically sorted."
Sorted = PropSorted
End Property

Public Property Let Sorted(ByVal Value As Boolean)
If PropVirtualMode = True And Value = True Then
    If ListViewDesignMode = True Then
        MsgBox "Sorted must be False when VirtualMode is True", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=383, Description:="Sorted must be False when VirtualMode is True"
    End If
End If
PropSorted = Value
If PropSorted = True And ListViewDesignMode = False Then Call SortListItems
UserControl.PropertyChanged "Sorted"
End Property

Public Property Get SortKey() As Integer
Attribute SortKey.VB_Description = "Returns/sets the current sort key."
SortKey = PropSortKey
End Property

Public Property Let SortKey(ByVal Value As Integer)
If Value < 0 Then
    If ListViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropSortKey = Value
If PropSorted = True And ListViewDesignMode = False Then Call SortListItems
UserControl.PropertyChanged "SortKey"
End Property

Public Property Get SortOrder() As LvwSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets a value that determines whether the list items will be sorted in ascending or descending order."
SortOrder = PropSortOrder
End Property

Public Property Let SortOrder(ByVal Value As LvwSortOrderConstants)
Select Case Value
    Case LvwSortOrderAscending, LvwSortOrderDescending
        PropSortOrder = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True And ListViewDesignMode = False Then Call SortListItems
UserControl.PropertyChanged "SortOrder"
End Property

Public Property Get SortType() As LvwSortTypeConstants
Attribute SortType.VB_Description = "Returns/sets the sort type."
SortType = PropSortType
End Property

Public Property Let SortType(ByVal Value As LvwSortTypeConstants)
Select Case Value
    Case LvwSortTypeBinary, LvwSortTypeText, LvwSortTypeNumeric, LvwSortTypeCurrency, LvwSortTypeDate, LvwSortTypeLogical
        PropSortType = Value
    Case Else
        Err.Raise 380
End Select
If PropSorted = True And ListViewDesignMode = False Then Call SortListItems
UserControl.PropertyChanged "SortType"
End Property

Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value that determines whether or not a checkbox is displayed next to each list item."
Checkboxes = PropCheckboxes
End Property

Public Property Let Checkboxes(ByVal Value As Boolean)
PropCheckboxes = Value
If ListViewHandle <> 0 Then
    If PropCheckboxes = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_CHECKBOXES, ByVal LVS_EX_CHECKBOXES
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_CHECKBOXES, ByVal 0&
    End If
End If
UserControl.PropertyChanged "Checkboxes"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value that determines whether the selected item will display as selected when the list view loses focus or not."
HideSelection = PropHideSelection
End Property

Public Property Let HideSelection(ByVal Value As Boolean)
PropHideSelection = Value
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If PropHideSelection = True Then
        If (dwStyle And LVS_SHOWSELALWAYS) = LVS_SHOWSELALWAYS Then dwStyle = dwStyle And Not LVS_SHOWSELALWAYS
    Else
        If Not (dwStyle And LVS_SHOWSELALWAYS) = LVS_SHOWSELALWAYS Then dwStyle = dwStyle Or LVS_SHOWSELALWAYS
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
    Me.Refresh
End If
UserControl.PropertyChanged "HideSelection"
End Property

Public Property Get HideColumnHeaders() As Boolean
Attribute HideColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the column headers are hidden in 'report' view."
HideColumnHeaders = PropHideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal Value As Boolean)
PropHideColumnHeaders = Value
If ListViewHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(ListViewHandle, GWL_STYLE)
    If PropHideColumnHeaders = True Then
        If Not (dwStyle And LVS_NOCOLUMNHEADER) = LVS_NOCOLUMNHEADER Then dwStyle = dwStyle Or LVS_NOCOLUMNHEADER
    Else
        If (dwStyle And LVS_NOCOLUMNHEADER) = LVS_NOCOLUMNHEADER Then dwStyle = dwStyle And Not LVS_NOCOLUMNHEADER
    End If
    SetWindowLong ListViewHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "HideColumnHeaders"
End Property

Public Property Get ShowInfoTips() As Boolean
Attribute ShowInfoTips.VB_Description = "Returns/sets a value that determines whether the tool tip text properties will be displayed or not."
ShowInfoTips = PropShowInfoTips
End Property

Public Property Let ShowInfoTips(ByVal Value As Boolean)
PropShowInfoTips = Value
If ListViewHandle <> 0 Then
    If PropShowInfoTips = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_INFOTIP, ByVal LVS_EX_INFOTIP
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_INFOTIP, ByVal 0&
    End If
End If
UserControl.PropertyChanged "ShowInfoTips"
End Property

Public Property Get ShowLabelTips() As Boolean
Attribute ShowLabelTips.VB_Description = "Returns/sets a value indicating that if a partially hidden label in any list view mode lacks tool tip text, the list view will unfold the label or not. Unfolding partially hidden labels for the 'icon' view are always done."
ShowLabelTips = PropShowLabelTips
End Property

Public Property Let ShowLabelTips(ByVal Value As Boolean)
PropShowLabelTips = Value
If ListViewHandle <> 0 Then
    If PropShowLabelTips = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_LABELTIP, ByVal LVS_EX_LABELTIP
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_LABELTIP, ByVal 0&
    End If
End If
UserControl.PropertyChanged "ShowLabelTips"
End Property

Public Property Get ShowColumnTips() As Boolean
Attribute ShowColumnTips.VB_Description = "Returns/sets a value that determines whether the column header tool tip text properties will be displayed or not."
ShowColumnTips = PropShowColumnTips
End Property

Public Property Let ShowColumnTips(ByVal Value As Boolean)
PropShowColumnTips = Value
If ListViewHandle <> 0 And ListViewDesignMode = False Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        If PropShowColumnTips = False Then
            Call DestroyHeaderToolTip
        Else
            Call CreateHeaderToolTip
        End If
    End If
End If
UserControl.PropertyChanged "ShowColumnTips"
End Property

Public Property Get DoubleBuffer() As Boolean
Attribute DoubleBuffer.VB_Description = "Returns/sets a value that determines whether the control paints via double-buffering, which reduces flicker. Requires comctl32.dll version 6.0 or higher."
DoubleBuffer = PropDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal Value As Boolean)
PropDoubleBuffer = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If PropDoubleBuffer = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_DOUBLEBUFFER, ByVal LVS_EX_DOUBLEBUFFER
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_DOUBLEBUFFER, ByVal 0&
    End If
End If
UserControl.PropertyChanged "DoubleBuffer"
End Property

Public Property Get HoverSelection() As Boolean
Attribute HoverSelection.VB_Description = "Returns/sets a value that determines whether or not an list item is automatically selected when the cursor remains over the list item for a certain period of time."
HoverSelection = PropHoverSelection
End Property

Public Property Let HoverSelection(ByVal Value As Boolean)
PropHoverSelection = Value
If ListViewHandle <> 0 Then
    If PropHotTracking = True Then
        If PropHighlightHot = True Or PropUnderlineHot = True Then
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT, ByVal LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT
        Else
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT, ByVal LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE
        End If
    Else
        If PropHoverSelection = True Then
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT, ByVal LVS_EX_TRACKSELECT
        Else
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT, ByVal 0&
        End If
    End If
End If
UserControl.PropertyChanged "HoverSelection"
End Property

Public Property Get HoverSelectionTime() As Long
Attribute HoverSelectionTime.VB_Description = "Returns/sets the hover selection time in milliseconds. A value of -1 indicates that the default time is used."
If ListViewHandle <> 0 Then
    HoverSelectionTime = SendMessage(ListViewHandle, LVM_GETHOVERTIME, 0, ByVal 0&)
Else
    HoverSelectionTime = PropHoverSelectionTime
End If
End Property

Public Property Let HoverSelectionTime(ByVal Value As Long)
If Value <= 0 And Not Value = -1 Then
    If ListViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropHoverSelectionTime = Value
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_SETHOVERTIME, 0, ByVal Value
UserControl.PropertyChanged "HoverSelectionTime"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets whether hot tracking is enabled."
HotTracking = PropHotTracking
End Property

Public Property Let HotTracking(ByVal Value As Boolean)
PropHotTracking = Value
If ListViewHandle <> 0 Then
    If PropHotTracking = True Then
        If PropHighlightHot = True Or PropUnderlineHot = True Then
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT, ByVal LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT
        Else
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT, ByVal LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE
        End If
    Else
        If PropHoverSelection = True Then
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT, ByVal LVS_EX_TRACKSELECT
        Else
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_TRACKSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT, ByVal 0&
        End If
    End If
End If
UserControl.PropertyChanged "HotTracking"
End Property

Public Property Get HighlightHot() As Boolean
Attribute HighlightHot.VB_Description = "Returns/sets a value that determines whether hot items that may be activated to be displayed with highlighted text. Only applicable if the hot tracking property is set to true."
HighlightHot = PropHighlightHot
End Property

Public Property Let HighlightHot(ByVal Value As Boolean)
PropHighlightHot = Value
If ListViewHandle <> 0 Then
    If PropHotTracking = True Then
        If PropHighlightHot = True Or PropUnderlineHot = True Then
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_UNDERLINEHOT, ByVal LVS_EX_UNDERLINEHOT
        Else
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_UNDERLINEHOT, ByVal 0&
        End If
    End If
End If
UserControl.PropertyChanged "HighlightHot"
End Property

Public Property Get UnderlineHot() As Boolean
Attribute UnderlineHot.VB_Description = "Returns/sets a value that determines whether hot items that may be activated to be displayed with underlined text or not. Only applicable if the hot tracking property is set to true."
UnderlineHot = PropUnderlineHot
End Property

Public Property Let UnderlineHot(ByVal Value As Boolean)
PropUnderlineHot = Value
If ListViewHandle <> 0 Then
    If PropHotTracking = True Then
        If PropHighlightHot = True Or PropUnderlineHot = True Then
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_UNDERLINEHOT, ByVal LVS_EX_UNDERLINEHOT
        Else
            SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_UNDERLINEHOT, ByVal 0&
        End If
    End If
End If
UserControl.PropertyChanged "UnderlineHot"
End Property

Public Property Get InsertMarkColor() As OLE_COLOR
Attribute InsertMarkColor.VB_Description = "Returns/sets the color of the insertion mark. Requires comctl32.dll version 6.1 or higher."
InsertMarkColor = PropInsertMarkColor
End Property

Public Property Let InsertMarkColor(ByVal Value As OLE_COLOR)
PropInsertMarkColor = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_SETINSERTMARKCOLOR, 0, ByVal WinColor(PropInsertMarkColor)
UserControl.PropertyChanged "InsertMarkColor"
End Property

Public Property Get TextBackground() As CCBackStyleConstants
Attribute TextBackground.VB_Description = "Returns/sets a value that determines if the text background is transparent or uses the background color of the list view."
TextBackground = PropTextBackground
End Property

Public Property Let TextBackground(ByVal Value As CCBackStyleConstants)
Select Case Value
    Case CCBackStyleTransparent, CCBackStyleOpaque
        PropTextBackground = Value
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 Then
    If PropTextBackground = CCBackStyleTransparent Then
        SendMessage ListViewHandle, LVM_SETTEXTBKCOLOR, 0, ByVal CLR_NONE
    Else
        SendMessage ListViewHandle, LVM_SETTEXTBKCOLOR, 0, ByVal WinColor(PropBackColor)
    End If
End If
UserControl.PropertyChanged "TextBackground"
End Property

Public Property Get ClickableColumnHeaders() As Boolean
Attribute ClickableColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the column headers act like buttons and are clickable in 'report' view."
ClickableColumnHeaders = PropClickableColumnHeaders
End Property

Public Property Let ClickableColumnHeaders(ByVal Value As Boolean)
PropClickableColumnHeaders = Value
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropClickableColumnHeaders = CBool((dwStyle And HDS_BUTTONS) = HDS_BUTTONS) Then
            If PropClickableColumnHeaders = True Then
                If Not (dwStyle And HDS_BUTTONS) = HDS_BUTTONS Then dwStyle = dwStyle Or HDS_BUTTONS
            Else
                If (dwStyle And HDS_BUTTONS) = HDS_BUTTONS Then dwStyle = dwStyle And Not HDS_BUTTONS
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "ClickableColumnHeaders"
End Property

Public Property Get HighlightColumnHeaders() As Boolean
Attribute HighlightColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the control highlights the column headers as the pointer passes over them. This flag is ignored on Windows XP (or above) when the desktop theme overrides it."
HighlightColumnHeaders = PropHighlightColumnHeaders
End Property

Public Property Let HighlightColumnHeaders(ByVal Value As Boolean)
PropHighlightColumnHeaders = Value
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropHighlightColumnHeaders = CBool((dwStyle And HDS_HOTTRACK) = HDS_HOTTRACK) Then
            If PropHighlightColumnHeaders = True Then
                If Not (dwStyle And HDS_HOTTRACK) = HDS_HOTTRACK Then dwStyle = dwStyle Or HDS_HOTTRACK
            Else
                If (dwStyle And HDS_HOTTRACK) = HDS_HOTTRACK Then dwStyle = dwStyle And Not HDS_HOTTRACK
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "HighlightColumnHeaders"
End Property

Public Property Get TrackSizeColumnHeaders() As Boolean
Attribute TrackSizeColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the control display column header contents even while the user resizes them."
TrackSizeColumnHeaders = PropTrackSizeColumnHeaders
End Property

Public Property Let TrackSizeColumnHeaders(ByVal Value As Boolean)
PropTrackSizeColumnHeaders = Value
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropTrackSizeColumnHeaders = CBool((dwStyle And HDS_FULLDRAG) = HDS_FULLDRAG) Then
            If PropTrackSizeColumnHeaders = True Then
                If Not (dwStyle And HDS_FULLDRAG) = HDS_FULLDRAG Then dwStyle = dwStyle Or HDS_FULLDRAG
            Else
                If (dwStyle And HDS_FULLDRAG) = HDS_FULLDRAG Then dwStyle = dwStyle And Not HDS_FULLDRAG
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "TrackSizeColumnHeaders"
End Property

Public Property Get ResizableColumnHeaders() As Boolean
Attribute ResizableColumnHeaders.VB_Description = "Returns/sets a value that determines whether or not the user can drag the divider on the column headers in 'report' view to resize them."
ResizableColumnHeaders = PropResizableColumnHeaders
End Property

Public Property Let ResizableColumnHeaders(ByVal Value As Boolean)
PropResizableColumnHeaders = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropResizableColumnHeaders = Not CBool((dwStyle And HDS_NOSIZING) = HDS_NOSIZING) Then
            If PropResizableColumnHeaders = True Then
                If (dwStyle And HDS_NOSIZING) = HDS_NOSIZING Then dwStyle = dwStyle And Not HDS_NOSIZING
            Else
                If Not (dwStyle And HDS_NOSIZING) = HDS_NOSIZING Then dwStyle = dwStyle Or HDS_NOSIZING
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "ResizableColumnHeaders"
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets the background picture. Requires comctl32.dll version 6.0 or higher."
Set Picture = PropPicture
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
Set Me.Picture = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
Dim LVBKI As LVBKIMAGE
With LVBKI
If Value Is Nothing Then
    Set PropPicture = Nothing
    If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
        .hBmp = 0
        .ulFlags = LVBKIF_SOURCE_NONE
        SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
        .ulFlags = LVBKIF_TYPE_WATERMARK
        SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
    End If
Else
    Set UserControl.Picture = Value
    Set PropPicture = UserControl.Picture
    Set UserControl.Picture = Nothing
    If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
        .hBmp = 0
        .ulFlags = LVBKIF_SOURCE_NONE
        SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
        .ulFlags = LVBKIF_TYPE_WATERMARK
        SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
        .ulFlags = LVBKIF_STYLE_NORMAL
        If Value.Handle <> 0 Then
            .hBmp = BitmapHandleFromPicture(PropPicture, PropBackColor)
            If PropPictureWatermark = False Then
                ' There is a much better result without LVS_EX_DOUBLEBUFFER
                ' when loading picture by 'hBmp'. (Weighing the pros and cons)
                If PropDoubleBuffer = True Then Me.DoubleBuffer = False
                Select Case PropPictureAlignment
                    Case LvwPictureAlignmentTopLeft
                        .XOffsetPercent = 0
                        .YOffsetPercent = 0
                    Case LvwPictureAlignmentTopRight
                        .XOffsetPercent = 100
                        .YOffsetPercent = 0
                    Case LvwPictureAlignmentBottomLeft
                        .XOffsetPercent = 0
                        .YOffsetPercent = 100
                    Case LvwPictureAlignmentBottomRight
                        .XOffsetPercent = 100
                        .YOffsetPercent = 100
                    Case LvwPictureAlignmentCenter
                        .XOffsetPercent = 50
                        .YOffsetPercent = 50
                    Case LvwPictureAlignmentTile
                        ' There is a better result when no column is selected.
                        Set Me.SelectedColumn = Nothing
                        .ulFlags = .ulFlags Or LVBKIF_STYLE_TILE
                        If ComCtlsSupportLevel() >= 2 And PropView = LvwViewReport Then
                            If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
                            If ListViewHeaderHandle <> 0 Then
                                .ulFlags = .ulFlags Or LVBKIF_FLAG_TILEOFFSET
                                Dim RC As RECT
                                GetWindowRect ListViewHeaderHandle, RC
                                .YOffsetPercent = -(RC.Bottom - RC.Top)
                            End If
                        End If
                End Select
                .ulFlags = .ulFlags Or LVBKIF_SOURCE_HBITMAP
            Else
                ' Here it does not matter whether LVS_EX_DOUBLEBUFFER is set or not.
                ' Though it is better to set it as it reduces flicker, especially
                ' when a watermark is in place.
                If PropDoubleBuffer = False Then Me.DoubleBuffer = True
                .ulFlags = .ulFlags Or LVBKIF_TYPE_WATERMARK
            End If
            SendMessage ListViewHandle, LVM_SETBKIMAGE, 0, ByVal VarPtr(LVBKI)
        End If
    End If
End If
End With
UserControl.PropertyChanged "Picture"
End Property

Public Property Get PictureAlignment() As LvwPictureAlignmentConstants
Attribute PictureAlignment.VB_Description = "Returns/sets the picture alignment. Requires comctl32.dll version 6.0 or higher."
PictureAlignment = PropPictureAlignment
End Property

Public Property Let PictureAlignment(ByVal Value As LvwPictureAlignmentConstants)
Select Case Value
    Case LvwPictureAlignmentTopLeft, LvwPictureAlignmentTopRight, LvwPictureAlignmentBottomLeft, LvwPictureAlignmentBottomRight, LvwPictureAlignmentCenter, LvwPictureAlignmentTile
        PropPictureAlignment = Value
    Case Else
        Err.Raise 380
End Select
Set Me.Picture = PropPicture
UserControl.PropertyChanged "PictureAlignment"
End Property

Public Property Get PictureWatermark() As Boolean
Attribute PictureWatermark.VB_Description = "Returns/sets a value that determines whether a watermark background bitmap is supplied in the picture property. That means the picture will always be displayed in the lower right corner. Requires comctl32.dll version 6.0 or higher."
PictureWatermark = PropPictureWatermark
End Property

Public Property Let PictureWatermark(ByVal Value As Boolean)
PropPictureWatermark = Value
Set Me.Picture = PropPicture
UserControl.PropertyChanged "PictureWatermark"
End Property

Public Property Get TileViewLines() As Long
Attribute TileViewLines.VB_Description = "Returns/sets the maximum number of text lines (not counting the title) in each list item in 'tile' view. Requires comctl32.dll version 6.0 or higher."
TileViewLines = PropTileViewLines
End Property

Public Property Let TileViewLines(ByVal Value As Long)
Select Case Value
    Case 0 To 20
        PropTileViewLines = Value
    Case Else
        If ListViewDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
End Select
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVTVI As LVTILEVIEWINFO
    With LVTVI
    .cbSize = LenB(LVTVI)
    .dwMask = LVTVIM_COLUMNS
    SendMessage ListViewHandle, LVM_GETTILEVIEWINFO, 0, ByVal VarPtr(LVTVI)
    .cLines = Value
    SendMessage ListViewHandle, LVM_SETTILEVIEWINFO, 0, ByVal VarPtr(LVTVI)
    End With
End If
UserControl.PropertyChanged "TileViewLines"
End Property

Public Property Get SnapToGrid() As Boolean
Attribute SnapToGrid.VB_Description = "Returns/sets a value that determines whether or not the list items automatically snaps into a grid in 'icon', 'small icon' or 'tile' view. Requires comctl32.dll version 6.0 or higher."
SnapToGrid = PropSnapToGrid
End Property

Public Property Let SnapToGrid(ByVal Value As Boolean)
PropSnapToGrid = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If PropSnapToGrid = True Then
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SNAPTOGRID, ByVal LVS_EX_SNAPTOGRID
    Else
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SNAPTOGRID, ByVal 0&
    End If
End If
UserControl.PropertyChanged "SnapToGrid"
End Property

Public Property Get GroupView() As Boolean
Attribute GroupView.VB_Description = "Returns/sets a value that determines whether or not the list items display as a group. Requires comctl32.dll version 6.0 or higher."
GroupView = PropGroupView
End Property

Public Property Let GroupView(ByVal Value As Boolean)
If PropVirtualMode = True And Value = True Then
    If ListViewDesignMode = True Then
        MsgBox "GroupView must be False when VirtualMode is True", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=383, Description:="GroupView must be False when VirtualMode is True"
    End If
End If
PropGroupView = Value
If ListViewDesignMode = False Then
    If ComCtlsSupportLevel() >= 1 Then
        If ListViewHandle <> 0 Then
            SendMessage ListViewHandle, LVM_ENABLEGROUPVIEW, IIf(PropGroupView = True, 1, 0), ByVal 0&
            Me.Refresh
        End If
    Else
        PropGroupView = False
    End If
End If
UserControl.PropertyChanged "GroupView"
End Property

Public Property Get GroupSubsetCount() As Long
Attribute GroupSubsetCount.VB_Description = "Returns/sets the number of list items that will be displayed in a subseted group. A value of 0 indicates that all list items are displayed, which means no subset. Requires comctl32.dll version 6.1 or higher."
GroupSubsetCount = PropGroupSubsetCount
End Property

Public Property Let GroupSubsetCount(ByVal Value As Long)
If Value < 0 Then
    If ListViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropGroupSubsetCount = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_SETGROUPSUBSETCOUNT, 0, ByVal PropGroupSubsetCount
UserControl.PropertyChanged "GroupSubsetCount"
End Property

Public Property Get UseColumnChevron() As Boolean
Attribute UseColumnChevron.VB_Description = "Returns/sets a value indicating if a chevron button is used when the column headers are wider than the control width. Requires comctl32.dll version 6.1 or higher."
UseColumnChevron = PropUseColumnChevron
End Property

Public Property Let UseColumnChevron(ByVal Value As Boolean)
PropUseColumnChevron = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropUseColumnChevron = CBool((dwStyle And HDS_OVERFLOW) = HDS_OVERFLOW) Then
            If PropUseColumnChevron = True Then
                If Not (dwStyle And HDS_OVERFLOW) = HDS_OVERFLOW Then dwStyle = dwStyle Or HDS_OVERFLOW
            Else
                If (dwStyle And HDS_OVERFLOW) = HDS_OVERFLOW Then dwStyle = dwStyle And Not HDS_OVERFLOW
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
        End If
    End If
End If
UserControl.PropertyChanged "UseColumnChevron"
End Property

Public Property Get UseColumnFilterBar() As Boolean
Attribute UseColumnFilterBar.VB_Description = "Returns/sets a value indicating if a filter bar is used on the column headers to allow users to conveniently apply a filter to the display."
UseColumnFilterBar = PropUseColumnFilterBar
End Property

Public Property Let UseColumnFilterBar(ByVal Value As Boolean)
PropUseColumnFilterBar = Value
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(ListViewHeaderHandle, GWL_STYLE)
        If Not PropUseColumnFilterBar = CBool((dwStyle And HDS_FILTERBAR) = HDS_FILTERBAR) Then
            If PropUseColumnFilterBar = True Then
                If Not (dwStyle And HDS_FILTERBAR) = HDS_FILTERBAR Then dwStyle = dwStyle Or HDS_FILTERBAR
            Else
                If (dwStyle And HDS_FILTERBAR) = HDS_FILTERBAR Then dwStyle = dwStyle And Not HDS_FILTERBAR
            End If
            SetWindowLong ListViewHeaderHandle, GWL_STYLE, dwStyle
            ' The header layout needs to be adjusted.
            Dim HDL As HDLAYOUT, RC As RECT, WPOS As WINDOWPOS
            GetClientRect ListViewHandle, RC
            HDL.lpRC = VarPtr(RC)
            HDL.lpWPOS = VarPtr(WPOS)
            SendMessage ListViewHeaderHandle, HDM_LAYOUT, 0, ByVal VarPtr(HDL)
            SetWindowPos WPOS.hWnd, WPOS.hWndInsertAfter, WPOS.X, WPOS.Y, WPOS.CX, WPOS.CY, WPOS.Flags
            ' Hide and show will force the necessary updates in the view area.
            ShowWindow ListViewHandle, SW_HIDE
            ShowWindow ListViewHandle, SW_SHOW
        End If
    End If
End If
UserControl.PropertyChanged "UseColumnFilterBar"
End Property

Public Property Get AutoSelectFirstItem() As Boolean
Attribute AutoSelectFirstItem.VB_Description = "Returns/sets a value that determines whether or not the first item will be selected automatically."
AutoSelectFirstItem = PropAutoSelectFirstItem
End Property

Public Property Let AutoSelectFirstItem(ByVal Value As Boolean)
PropAutoSelectFirstItem = Value
UserControl.PropertyChanged "AutoSelectFirstItem"
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
If ListViewHandle <> 0 And ListViewDesignMode = False Then
    If GetFocus() = ListViewHandle Then
        Call ComCtlsSetIMEMode(ListViewHandle, ListViewIMCHandle, PropIMEMode)
    ElseIf ListViewFilterEditHandle <> 0 Then
        If GetFocus() = ListViewFilterEditHandle Then Call ComCtlsSetIMEMode(ListViewFilterEditHandle, ListViewIMCHandle, PropIMEMode)
    ElseIf ListViewLabelInEdit = True Then
        Dim LabelEditHandle As Long
        LabelEditHandle = Me.hWndLabelEdit
        If LabelEditHandle <> 0 Then
            If GetFocus() = LabelEditHandle Then Call ComCtlsSetIMEMode(LabelEditHandle, ListViewIMCHandle, PropIMEMode)
        End If
    End If
End If
UserControl.PropertyChanged "IMEMode"
End Property

Public Property Get VirtualMode() As Boolean
Attribute VirtualMode.VB_Description = "Returns/sets a value indicating if you have provided your own data-management operations for the control."
VirtualMode = PropVirtualMode
End Property

Public Property Let VirtualMode(ByVal Value As Boolean)
If ListViewDesignMode = False Then
    Err.Raise Number:=382, Description:="VirtualMode property is read-only at run time"
Else
    PropVirtualMode = Value
    If PropVirtualMode = True Then
        If PropView = LvwViewTile Then PropView = LvwViewIcon
        PropArrange = LvwArrangeNone
        PropSorted = False
        PropGroupView = False
    End If
End If
UserControl.PropertyChanged "VirtualMode"
End Property

Public Property Get VirtualItemCount() As Long
Attribute VirtualItemCount.VB_Description = "Returns/sets the virtual number of items that the control contains."
VirtualItemCount = PropVirtualItemCount
End Property

Public Property Let VirtualItemCount(ByVal Value As Long)
If Value < 0 Or Value > 100000000 Then
    ' According to MSDN:
    ' There is a 100,000,000 item limit on a virtualized list view.
    If ListViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If PropVirtualMode = True Then
    If ListViewHandle <> 0 And ListViewDesignMode = False Then
        If SendMessage(ListViewHandle, LVM_SETITEMCOUNT, Value, ByVal 0&) = 0 Then Err.Raise 380
        If ListViewListItemsControl = 0 Then
            Dim LVI As LVITEM
            With LVI
            If PropAutoSelectFirstItem = True Then
                .StateMask = LVIS_SELECTED Or LVIS_FOCUSED
                .State = LVIS_SELECTED Or LVIS_FOCUSED
            Else
                .StateMask = LVIS_FOCUSED
                .State = LVIS_FOCUSED
            End If
            End With
            SendMessage ListViewHandle, LVM_SETITEMSTATE, 0, ByVal VarPtr(LVI)
        End If
        ListViewListItemsControl = Value
    End If
End If
PropVirtualItemCount = Value
UserControl.PropertyChanged "VirtualItemCount"
End Property

Public Property Get VirtualDisabledInfos() As LvwVirtualPropertyConstants
Attribute VirtualDisabledInfos.VB_Description = "Returns/sets the disabled virtual properties that are not needed and to increase performance."
Attribute VirtualDisabledInfos.VB_MemberFlags = "400"
VirtualDisabledInfos = PropVirtualDisabledInfos
End Property

Public Property Let VirtualDisabledInfos(ByVal Value As LvwVirtualPropertyConstants)
If Value < 0 Then
    If ListViewDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
PropVirtualDisabledInfos = Value
UserControl.PropertyChanged "VirtualDisabledInfos"
End Property

Public Property Get ListItems() As LvwListItems
Attribute ListItems.VB_Description = "Returns a reference to a collection of the list item objects."
If PropListItems Is Nothing Then
    If PropVirtualMode = False Then
        Set PropListItems = New LvwListItems
        PropListItems.FInit Me
    Else
        Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
    End If
End If
Set ListItems = PropListItems
End Property

Public Property Get VirtualListItems() As LvwVirtualListItems
Attribute VirtualListItems.VB_Description = "Returns a reference to a collection of the virtual list item objects."
If PropVirtualMode = False Then Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is off."
Set VirtualListItems = New LvwVirtualListItems
VirtualListItems.FInit Me
End Property

Friend Sub FListItemsAdd(ByVal Ptr As Long, ByVal Index As Long, Optional ByVal Text As String)
Dim LVI As LVITEM
With LVI
.Mask = LVIF_TEXT Or LVIF_IMAGE Or LVIF_PARAM Or LVIF_INDENT
.iItem = Index - 1
.pszText = LPSTR_TEXTCALLBACK
.iImage = I_IMAGECALLBACK
.lParam = Ptr
.iIndent = 0
End With
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_INSERTITEM, 0, ByVal VarPtr(LVI)
If PropSorted = True Then If PropSortKey = 0 Then Call SortListItems
End Sub

Friend Sub FListItemsRemove(ByVal Index As Long)
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, LVM_DELETEITEM, Index - 1, ByVal 0&
    If ListViewListItemsControl = 0 Then
        Call CheckItemFocus(0)
    ElseIf ListViewFocusIndex > Index Then
        ListViewFocusIndex = ListViewFocusIndex - 1
    End If
End If
End Sub

Friend Sub FListItemsClear()
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_DELETEALLITEMS, 0, ByVal 0&
Call CheckItemFocus(0)
End Sub

Friend Function FListItemPtr(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    .Mask = LVIF_PARAM
    .iItem = Index - 1
    SendMessage ListViewHandle, LVM_GETITEM, 0, ByVal VarPtr(LVI)
    FListItemPtr = .lParam
    End With
End If
End Function

Friend Function FListItemVerify(ByVal Ptr As Long, ByRef Index As Long) As Boolean
If Ptr = Me.FListItemPtr(Index) Or Ptr = 0 Then
    FListItemVerify = True
Else
    Index = Me.FListItemIndex(Ptr)
    FListItemVerify = CBool(Index <> 0)
End If
End Function

Friend Function FListItemIndex(ByVal Ptr As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVFI As LVFINDINFO
    With LVFI
    .Flags = LVFI_PARAM
    .lParam = Ptr
    End With
    FListItemIndex = SendMessage(ListViewHandle, LVM_FINDITEM, -1, ByVal VarPtr(LVFI)) + 1
End If
End Function

Friend Sub FListItemRedraw(ByVal Index As Long)
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, LVM_REDRAWITEMS, Index - 1, ByVal (Index - 1)
    UpdateWindow ListViewHandle
End If
End Sub

Friend Property Get FListItemText(ByVal Index As Long, ByVal SubItemIndex As Long) As String
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    Dim Buffer As String
    Buffer = String(260, vbNullChar)
    .pszText = StrPtr(Buffer)
    .cchTextMax = 260
    .iSubItem = SubItemIndex
    End With
    SendMessage ListViewHandle, LVM_GETITEMTEXT, Index - 1, ByVal VarPtr(LVI)
    FListItemText = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Friend Property Let FListItemText(ByVal Index As Long, ByVal SubItemIndex As Long, ByVal Value As String)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    .pszText = LPSTR_TEXTCALLBACK
    .iSubItem = SubItemIndex
    End With
    SendMessage ListViewHandle, LVM_SETITEMTEXT, Index - 1, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemIndentation(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    LVI.Mask = LVIF_INDENT
    LVI.iItem = Index - 1
    SendMessage ListViewHandle, LVM_GETITEM, 0, ByVal VarPtr(LVI)
    FListItemIndentation = LVI.iIndent
End If
End Property

Friend Property Let FListItemIndentation(ByVal Index As Long, ByVal Value As Long)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    LVI.Mask = LVIF_INDENT
    LVI.iItem = Index - 1
    LVI.iIndent = Value
    SendMessage ListViewHandle, LVM_SETITEM, 0, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemSelected(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then FListItemSelected = CBool((SendMessage(ListViewHandle, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_SELECTED) And LVIS_SELECTED) = LVIS_SELECTED)
End Property

Friend Property Let FListItemSelected(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    If Value = True Then
        .StateMask = LVIS_SELECTED Or LVIS_FOCUSED
        .State = LVIS_SELECTED Or LVIS_FOCUSED
    Else
        .StateMask = LVIS_SELECTED
        .State = 0
    End If
    End With
    SendMessage ListViewHandle, LVM_SETITEMSTATE, Index - 1, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemChecked(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then FListItemChecked = CBool(StateImageMaskToIndex(SendMessage(ListViewHandle, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_STATEIMAGEMASK) And LVIS_STATEIMAGEMASK) = IIL_CHECKED)
End Property

Friend Property Let FListItemChecked(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    .StateMask = LVIS_STATEIMAGEMASK
    If Value = True Then
        .State = IndexToStateImageMask(IIL_CHECKED)
    Else
        .State = IndexToStateImageMask(IIL_UNCHECKED)
    End If
    End With
    SendMessage ListViewHandle, LVM_SETITEMSTATE, Index - 1, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemGhosted(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then FListItemGhosted = CBool((SendMessage(ListViewHandle, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_CUT) And LVIS_CUT) = LVIS_CUT)
End Property

Friend Property Let FListItemGhosted(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    Dim LVI As LVITEM
    With LVI
    .StateMask = LVIS_CUT
    If Value = True Then
        .State = LVIS_CUT
    Else
        .State = 0
    End If
    End With
    SendMessage ListViewHandle, LVM_SETITEMSTATE, Index - 1, ByVal VarPtr(LVI)
End If
End Property

Friend Property Get FListItemHot(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETHOTITEM, 0, ByVal 0&)
    If iItem > -1 Then FListItemHot = CBool(Index = (iItem + 1))
End If
End Property

Friend Property Let FListItemHot(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    If Value = True Then
        SendMessage ListViewHandle, LVM_SETHOTITEM, Index - 1, ByVal 0&
    Else
        If SendMessage(ListViewHandle, LVM_GETHOTITEM, 0, ByVal 0&) = Index - 1 Then SendMessage ListViewHandle, LVM_SETHOTITEM, -1, ByVal 0&
    End If
End If
End Property

Friend Property Get FListItemLeft(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim P As POINTAPI
    SendMessage ListViewHandle, LVM_GETITEMPOSITION, Index - 1, ByVal VarPtr(P)
    FListItemLeft = UserControl.ScaleX(P.X, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Let FListItemLeft(ByVal Index As Long, ByVal Value As Single)
If ListViewHandle <> 0 Then
    Dim P As POINTAPI
    SendMessage ListViewHandle, LVM_GETITEMPOSITION, Index - 1, ByVal VarPtr(P)
    P.X = UserControl.ScaleX(Value, vbContainerPosition, vbPixels)
    SendMessage ListViewHandle, LVM_SETITEMPOSITION32, Index - 1, ByVal VarPtr(P)
End If
End Property

Friend Property Get FListItemTop(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim P As POINTAPI
    SendMessage ListViewHandle, LVM_GETITEMPOSITION, Index - 1, ByVal VarPtr(P)
    FListItemTop = UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Let FListItemTop(ByVal Index As Long, ByVal Value As Single)
If ListViewHandle <> 0 Then
    Dim P As POINTAPI
    SendMessage ListViewHandle, LVM_GETITEMPOSITION, Index - 1, ByVal VarPtr(P)
    P.Y = UserControl.ScaleY(Value, vbContainerPosition, vbPixels)
    SendMessage ListViewHandle, LVM_SETITEMPOSITION32, Index - 1, ByVal VarPtr(P)
End If
End Property

Friend Property Get FListItemWidth(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_SELECTBOUNDS
    SendMessage ListViewHandle, LVM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListItemWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FListItemHeight(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_SELECTBOUNDS
    SendMessage ListViewHandle, LVM_GETITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListItemHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FListItemVisible(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then FListItemVisible = CBool(SendMessage(ListViewHandle, LVM_ISITEMVISIBLE, Index - 1, ByVal 0&) <> 0)
End Property

Friend Sub FListItemEnsureVisible(ByVal Index As Long)
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_ENSUREVISIBLE, Index - 1, ByVal 0&
End Sub

Friend Property Get FListItemTileViewIndices(ByVal Index As Long) As Variant
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim Buffer(0 To 19) As Long
    Dim LVTI As LVTILEINFO
    With LVTI
    .cbSize = LenB(LVTI)
    .iItem = Index - 1
    .cColumns = 20
    .puColumns = VarPtr(Buffer(0))
    SendMessage ListViewHandle, LVM_GETTILEINFO, 0, ByVal VarPtr(LVTI)
    If .cColumns > 0 Then
        Dim ArgList() As Long, i As Long
        ReDim ArgList(0 To (.cColumns - 1)) As Long
        For i = 0 To (.cColumns - 1)
            ArgList(i) = Buffer(i)
        Next i
        FListItemTileViewIndices = ArgList()
    Else
        FListItemTileViewIndices = Empty
    End If
    End With
End If
End Property

Friend Property Let FListItemTileViewIndices(ByVal Index As Long, ByVal ArgList As Variant)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVTI As LVTILEINFO
    With LVTI
    .cbSize = LenB(LVTI)
    .iItem = Index - 1
    If IsArray(ArgList) Then
        Dim Ptr As Long
        CopyMemory Ptr, ByVal UnsignedAdd(VarPtr(ArgList), 8), 4
        If Ptr <> 0 Then
            Dim DimensionCount As Integer
            CopyMemory DimensionCount, ByVal Ptr, 2
            If DimensionCount = 1 Then
                Dim Arr() As Long, Count As Long, i As Long
                For i = LBound(ArgList) To UBound(ArgList)
                    Select Case VarType(ArgList(i))
                        Case vbLong, vbInteger, vbByte
                            If ArgList(i) > 0 Then
                                ReDim Preserve Arr(0 To Count) As Long
                                Arr(Count) = ArgList(i)
                                Count = Count + 1
                            End If
                        Case vbDouble, vbSingle
                            If CLng(ArgList(i)) > 0 Then
                                ReDim Preserve Arr(0 To Count) As Long
                                Arr(Count) = CLng(ArgList(i))
                                Count = Count + 1
                            End If
                    End Select
                Next i
                If Count > 0 Then
                    .cColumns = Count
                    .puColumns = VarPtr(Arr(0))
                Else
                    .cColumns = 0
                    .puColumns = 0
                End If
            Else
                Err.Raise Number:=5, Description:="Array must be single dimensioned"
            End If
        Else
            Err.Raise Number:=91, Description:="Array is not allocated"
        End If
    ElseIf IsEmpty(ArgList) Then
        .cColumns = 0
        .puColumns = 0
    Else
        Err.Raise 380
    End If
    SendMessage ListViewHandle, LVM_SETTILEINFO, 0, ByVal VarPtr(LVTI)
    End With
End If
End Property

Friend Function FListItemCreateDragImage(ByVal Index As Long, ByRef X As Single, ByRef Y As Single) As Long
If ListViewHandle <> 0 Then
    Dim P As POINTAPI
    FListItemCreateDragImage = SendMessage(ListViewHandle, LVM_CREATEDRAGIMAGE, Index - 1, ByVal VarPtr(P))
    X = UserControl.ScaleX(P.X, vbPixels, vbContainerPosition)
    Y = UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition)
End If
End Function

Friend Property Get FListItemGroup(ByVal Index As Long) As LvwGroup
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVI_V60 As LVITEM_V60
    With LVI_V60
    .LVI.Mask = LVIF_GROUPID
    .LVI.iItem = Index - 1
    SendMessage ListViewHandle, LVM_GETITEM, 0, ByVal VarPtr(LVI_V60)
    If .iGroupId <> I_GROUPIDNONE Then
        Dim Group As LvwGroup
        For Each Group In Me.Groups
            If Group.ID = .iGroupId Then
                Set FListItemGroup = Group
                Exit For
            End If
        Next Group
    End If
    End With
End If
End Property

Friend Property Let FListItemGroup(ByVal Index As Long, ByVal Value As LvwGroup)
Set Me.FListItemGroup(Index) = Value
End Property

Friend Property Set FListItemGroup(ByVal Index As Long, ByVal Value As LvwGroup)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVI_V60 As LVITEM_V60
    With LVI_V60
    .LVI.Mask = LVIF_GROUPID
    .LVI.iItem = Index - 1
    If Not Value Is Nothing Then
        .iGroupId = Value.ID
    Else
        .iGroupId = I_GROUPIDNONE
    End If
    End With
    SendMessage ListViewHandle, LVM_SETITEM, 0, ByVal VarPtr(LVI_V60)
End If
End Property

Friend Property Get FListItemWorkArea(ByVal Index As Long) As LvwWorkArea
Select Case PropView
    Case LvwViewIcon, LvwViewSmallIcon, LvwViewTile
        If ListViewHandle <> 0 Then
            Dim Count As Long
            SendMessage ListViewHandle, LVM_GETNUMBEROFWORKAREAS, 0, ByVal VarPtr(Count)
            If Count > 0 Then
                Dim P As POINTAPI
                If SendMessage(ListViewHandle, LVM_GETITEMPOSITION, Index - 1, ByVal VarPtr(P)) <> 0 Then
                    Dim ArrRC() As RECT, iWorkArea As Long
                    ReDim ArrRC(1 To Count) As RECT
                    SendMessage ListViewHandle, LVM_GETWORKAREAS, Count, ByVal VarPtr(ArrRC(1))
                    For iWorkArea = 1 To Count
                        If PtInRect(ArrRC(iWorkArea), P.X, P.Y) <> 0 Then
                            Set FListItemWorkArea = New LvwWorkArea
                            FListItemWorkArea.FInit ObjPtr(Me), iWorkArea
                            Exit For
                        End If
                    Next iWorkArea
                End If
            End If
        End If
    Case Else
        Err.Raise Number:=394, Description:="Get supported in 'icon', 'small icon' and 'tile' view only"
End Select
End Property

Friend Property Get FListSubItemLeft(ByVal Index As Long, ByVal SubItemIndex As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_BOUNDS
    RC.Top = SubItemIndex
    SendMessage ListViewHandle, LVM_GETSUBITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListSubItemLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Get FListSubItemTop(ByVal Index As Long, ByVal SubItemIndex As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_BOUNDS
    RC.Top = SubItemIndex
    SendMessage ListViewHandle, LVM_GETSUBITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListSubItemTop = UserControl.ScaleX(RC.Top, vbPixels, vbContainerPosition)
End If
End Property

Friend Property Get FListSubItemWidth(ByVal Index As Long, ByVal SubItemIndex As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_BOUNDS
    RC.Top = SubItemIndex
    SendMessage ListViewHandle, LVM_GETSUBITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListSubItemWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FListSubItemHeight(ByVal Index As Long, ByVal SubItemIndex As Long) As Single
If ListViewHandle <> 0 Then
    Dim RC As RECT
    RC.Left = LVIR_BOUNDS
    RC.Top = SubItemIndex
    SendMessage ListViewHandle, LVM_GETSUBITEMRECT, Index - 1, ByVal VarPtr(RC)
    FListSubItemHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End If
End Property

Public Property Get ColumnHeaders() As LvwColumnHeaders
Attribute ColumnHeaders.VB_Description = "Returns a reference to a collection of the column header objects."
If PropColumnHeaders Is Nothing Then
    Set PropColumnHeaders = New LvwColumnHeaders
    PropColumnHeaders.FInit Me
End If
Set ColumnHeaders = PropColumnHeaders
End Property

Friend Sub FColumnHeadersAdd(ByVal Index As Long, Optional ByVal Text As String, Optional ByVal Width As Single, Optional ByVal Alignment As LvwColumnHeaderAlignmentConstants, Optional ByVal IconIndex As Long)
Dim ColumnHeaderIndex As Long
If Index = 0 Then
    ColumnHeaderIndex = Me.ColumnHeaders.Count + 1
Else
    ColumnHeaderIndex = Index
End If
Dim LVC As LVCOLUMN
With LVC
.Mask = LVCF_FMT Or LVCF_WIDTH
If Not Text = vbNullString Then
    .Mask = .Mask Or LVCF_TEXT
    .pszText = StrPtr(Text)
    .cchTextMax = Len(Text) + 1
End If
If Width = 0 Then
    .CX = (96 * PixelsPerDIP_X())
ElseIf Width > 0 Then
    .CX = UserControl.ScaleX(Width, vbContainerSize, vbPixels)
Else
    Err.Raise 380
End If
If (ColumnHeaderIndex - 1) = 0 Then
    .fmt = LVCFMT_LEFT
Else
    Select Case Alignment
        Case LvwColumnHeaderAlignmentLeft
            .fmt = LVCFMT_LEFT
        Case LvwColumnHeaderAlignmentRight
            .fmt = LVCFMT_RIGHT
        Case LvwColumnHeaderAlignmentCenter
            .fmt = LVCFMT_CENTER
        Case Else
            Err.Raise 380
    End Select
End If
If IconIndex > 0 Then
    .fmt = .fmt Or LVCFMT_IMAGE
    .Mask = .Mask Or LVCF_IMAGE
    .iImage = IconIndex - 1
End If
End With
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
    SendMessage ListViewHandle, LVM_INSERTCOLUMN, ColumnHeaderIndex - 1, ByVal VarPtr(LVC)
    If (ColumnHeaderIndex - 1) = 0 Then
        ' According to MSDN:
        ' If a column is added to a list view control with index 0 (the leftmost column), it is always LVCFMT_LEFT.
        ' Workaround: Adjust the fmt value after the insert.
        If Alignment <> LvwColumnHeaderAlignmentLeft Then Me.FColumnHeaderAlignment(1) = Alignment
    End If
    Call SetColumnsSubItemIndex(1)
    If PropRightToLeft = True And PropRightToLeftLayout = False Then Call SetColumnRTLReading(ColumnHeaderIndex, True)
    Call RebuildListItems
    If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
End If
End Sub

Friend Sub FColumnHeadersRemove(ByVal Index As Long)
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
    SendMessage ListViewHandle, LVM_DELETECOLUMN, Index - 1, ByVal 0&
    Call SetColumnsSubItemIndex(-1)
    Call RebuildListItems
    If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
End If
End Sub

Friend Sub FColumnHeadersClear()
If ListViewHandle <> 0 Then Do While SendMessage(ListViewHandle, LVM_DELETECOLUMN, 0, ByVal 0&) = 1: Loop
End Sub

Friend Sub FColumnHeadersRedraw()
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        InvalidateRect ListViewHeaderHandle, ByVal 0&, 1
        UpdateWindow ListViewHeaderHandle
    End If
End If
End Sub

Friend Function FColumnHeadersPositionToIndex(ByVal Position As Long) As Long
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then FColumnHeadersPositionToIndex = SendMessage(ListViewHeaderHandle, HDM_ORDERTOINDEX, Position - 1, ByVal 0&) + 1
End If
End Function

Friend Property Get FColumnHeaderText(ByVal Index As Long) As String
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_TEXT
    Dim Buffer As String
    Buffer = String(260, vbNullChar)
    .pszText = StrPtr(Buffer)
    .cchTextMax = 260
    End With
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderText = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End If
End Property

Friend Property Let FColumnHeaderText(ByVal Index As Long, ByVal Value As String)
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_TEXT
    .pszText = StrPtr(Value)
    End With
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
End If
End Property

Friend Property Get FColumnHeaderIcon(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If (LVC.fmt And LVCFMT_IMAGE) = LVCFMT_IMAGE Then
        LVC.Mask = LVCF_IMAGE
        SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        FColumnHeaderIcon = LVC.iImage + 1
    End If
End If
End Property

Friend Property Let FColumnHeaderIcon(ByVal Index As Long, ByVal Value As Long)
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    .Mask = LVCF_FMT Or LVCF_IMAGE
    .iImage = Value - 1
    If Value > 0 Then
        If Not (.fmt And LVCFMT_IMAGE) = LVCFMT_IMAGE Then .fmt = .fmt Or LVCFMT_IMAGE
    Else
        If (.fmt And LVCFMT_IMAGE) = LVCFMT_IMAGE Then .fmt = .fmt And Not LVCFMT_IMAGE
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderWidth(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_WIDTH
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderWidth = UserControl.ScaleX(LVC.CX, vbPixels, vbContainerSize)
End If
End Property

Friend Property Let FColumnHeaderWidth(ByVal Index As Long, ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_WIDTH
    .CX = UserControl.ScaleX(Value, vbContainerSize, vbPixels)
    End With
    If PropView = LvwViewReport Then
        SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    Else
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
    End If
End If
End Property

Friend Property Get FColumnHeaderAlignment(ByVal Index As Long) As LvwColumnHeaderAlignmentConstants
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If (.fmt And LVCFMT_CENTER) = LVCFMT_CENTER Then
        FColumnHeaderAlignment = LvwColumnHeaderAlignmentCenter
    ElseIf (.fmt And LVCFMT_RIGHT) = LVCFMT_RIGHT Then
        FColumnHeaderAlignment = LvwColumnHeaderAlignmentRight
    ElseIf (.fmt And LVCFMT_LEFT) = LVCFMT_LEFT Then
        FColumnHeaderAlignment = LvwColumnHeaderAlignmentLeft
    End If
    End With
End If
End Property

Friend Property Let FColumnHeaderAlignment(ByVal Index As Long, ByVal Value As LvwColumnHeaderAlignmentConstants)
If ListViewHandle <> 0 Then
    Select Case Value
        Case LvwColumnHeaderAlignmentLeft, LvwColumnHeaderAlignmentRight, LvwColumnHeaderAlignmentCenter
            Dim LVC As LVCOLUMN
            With LVC
            .Mask = LVCF_FMT
            SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
            If (.fmt And LVCFMT_LEFT) = LVCFMT_LEFT Then .fmt = .fmt And Not LVCFMT_LEFT
            If (.fmt And LVCFMT_RIGHT) = LVCFMT_RIGHT Then .fmt = .fmt And Not LVCFMT_RIGHT
            If (.fmt And LVCFMT_CENTER) = LVCFMT_CENTER Then .fmt = .fmt And Not LVCFMT_CENTER
            Select Case Value
                Case LvwColumnHeaderAlignmentLeft
                    .fmt = .fmt Or LVCFMT_LEFT
                Case LvwColumnHeaderAlignmentRight
                    .fmt = .fmt Or LVCFMT_RIGHT
                Case LvwColumnHeaderAlignmentCenter
                    .fmt = .fmt Or LVCFMT_CENTER
            End Select
            End With
            SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        Case Else
            Err.Raise 380
    End Select
End If
End Property

Friend Property Get FColumnHeaderPosition(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_ORDER
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderPosition = LVC.iOrder + 1
End If
End Property

Friend Property Let FColumnHeaderPosition(ByVal Index As Long, ByVal Value As Long)
If ListViewHandle <> 0 Then
    If Value < 1 Or Value > Me.ColumnHeaders.Count Then
        Err.Raise 380
    Else
        Dim LVC As LVCOLUMN
        With LVC
        .Mask = LVCF_ORDER
        .iOrder = Value - 1
        End With
        SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        Me.Refresh
    End If
End If
End Property

Friend Property Get FColumnHeaderSortArrow(ByVal Index As Long) As LvwColumnHeaderSortArrowConstants
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If (.fmt And HDF_SORTUP) = HDF_SORTUP Then
        FColumnHeaderSortArrow = LvwColumnHeaderSortArrowUp
    ElseIf (.fmt And HDF_SORTDOWN) = HDF_SORTDOWN Then
        FColumnHeaderSortArrow = LvwColumnHeaderSortArrowDown
    Else
        FColumnHeaderSortArrow = LvwColumnHeaderSortArrowNone
    End If
    End With
End If
End Property

Friend Property Let FColumnHeaderSortArrow(ByVal Index As Long, ByVal Value As LvwColumnHeaderSortArrowConstants)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Select Case Value
        Case LvwColumnHeaderSortArrowNone, LvwColumnHeaderSortArrowDown, LvwColumnHeaderSortArrowUp
            Dim LVC As LVCOLUMN
            With LVC
            .Mask = LVCF_FMT
            SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
            If (.fmt And HDF_SORTDOWN) = HDF_SORTDOWN Then .fmt = .fmt And Not HDF_SORTDOWN
            If (.fmt And HDF_SORTUP) = HDF_SORTUP Then .fmt = .fmt And Not HDF_SORTUP
            Select Case Value
                Case LvwColumnHeaderSortArrowDown
                    .fmt = .fmt Or HDF_SORTDOWN
                Case LvwColumnHeaderSortArrowUp
                    .fmt = .fmt Or HDF_SORTUP
            End Select
            SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
            End With
        Case Else
            Err.Raise 380
    End Select
End If
End Property

Friend Property Get FColumnHeaderIconOnRight(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderIconOnRight = CBool((.fmt And LVCFMT_BITMAP_ON_RIGHT) = LVCFMT_BITMAP_ON_RIGHT)
    End With
End If
End Property

Friend Property Let FColumnHeaderIconOnRight(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        If Not (.fmt And LVCFMT_BITMAP_ON_RIGHT) = LVCFMT_BITMAP_ON_RIGHT Then .fmt = .fmt Or LVCFMT_BITMAP_ON_RIGHT
    Else
        If (.fmt And LVCFMT_BITMAP_ON_RIGHT) = LVCFMT_BITMAP_ON_RIGHT Then .fmt = .fmt And Not LVCFMT_BITMAP_ON_RIGHT
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderResizable(ByVal Index As Long, ByRef Resizable As Boolean) As Boolean
If ListViewHandle <> 0 Then
    If ComCtlsSupportLevel() >= 2 Then
        Dim LVC As LVCOLUMN
        With LVC
        .Mask = LVCF_FMT
        SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
        FColumnHeaderResizable = Not CBool((.fmt And LVCFMT_FIXED_WIDTH) = LVCFMT_FIXED_WIDTH)
        End With
    Else
        FColumnHeaderResizable = Resizable
    End If
End If
End Property

Friend Property Let FColumnHeaderResizable(ByVal Index As Long, ByRef Resizable As Boolean, ByVal Value As Boolean)
Resizable = Value
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        If (.fmt And LVCFMT_FIXED_WIDTH) = LVCFMT_FIXED_WIDTH Then .fmt = .fmt And Not LVCFMT_FIXED_WIDTH
    Else
        If Not (.fmt And LVCFMT_FIXED_WIDTH) = LVCFMT_FIXED_WIDTH Then .fmt = .fmt Or LVCFMT_FIXED_WIDTH
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderSplitButton(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderSplitButton = CBool((.fmt And LVCFMT_SPLITBUTTON) = LVCFMT_SPLITBUTTON)
    End With
End If
End Property

Friend Property Let FColumnHeaderSplitButton(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        If Not (.fmt And LVCFMT_SPLITBUTTON) = LVCFMT_SPLITBUTTON Then .fmt = .fmt Or LVCFMT_SPLITBUTTON
    Else
        If (.fmt And LVCFMT_SPLITBUTTON) = LVCFMT_SPLITBUTTON Then .fmt = .fmt And Not LVCFMT_SPLITBUTTON
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderCheckBox(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderCheckBox = CBool((.fmt And HDF_CHECKBOX) = HDF_CHECKBOX)
    End With
End If
End Property

Friend Property Let FColumnHeaderCheckBox(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If Value = True Then
        If Not (.fmt And HDF_CHECKBOX) = HDF_CHECKBOX Then .fmt = .fmt Or HDF_CHECKBOX
    Else
        If (.fmt And HDF_CHECKBOX) = HDF_CHECKBOX Then .fmt = .fmt And Not HDF_CHECKBOX
        If (.fmt And HDF_CHECKED) = HDF_CHECKED Then .fmt = .fmt And Not HDF_CHECKED
    End If
    SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    End With
End If
End Property

Friend Property Get FColumnHeaderChecked(ByVal Index As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderChecked = CBool((.fmt And HDF_CHECKED) = HDF_CHECKED)
    End With
End If
End Property

Friend Property Let FColumnHeaderChecked(ByVal Index As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVC As LVCOLUMN
    With LVC
    .Mask = LVCF_FMT
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    If (.fmt And HDF_CHECKBOX) = HDF_CHECKBOX Then
        If CBool((.fmt And HDF_CHECKED) = HDF_CHECKED) <> Value Then
            If Value = True Then
                If Not (.fmt And HDF_CHECKED) = HDF_CHECKED Then .fmt = .fmt Or HDF_CHECKED
            Else
                If (.fmt And HDF_CHECKED) = HDF_CHECKED Then .fmt = .fmt And Not HDF_CHECKED
            End If
            SendMessage ListViewHandle, LVM_SETCOLUMN, Index - 1, ByVal VarPtr(LVC)
            RaiseEvent ColumnCheck(Me.ColumnHeaders(Index))
        End If
    End If
    End With
End If
End Property

Friend Property Get FColumnHeaderFilterType(ByVal Index As Long) As LvwColumnHeaderFilterTypeConstants
If ListViewHandle <> 0 Then
    ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim HDI As HDITEM
        With HDI
        .Mask = HDI_FILTER
        SendMessage ListViewHeaderHandle, HDM_GETITEM, Index - 1, ByVal VarPtr(HDI)
        If (.FilterType And HDFT_HASNOVALUE) = HDFT_HASNOVALUE Then .FilterType = .FilterType And Not HDFT_HASNOVALUE
        Select Case .FilterType
            Case HDFT_ISSTRING
                FColumnHeaderFilterType = LvwColumnHeaderFilterTypeText
            Case HDFT_ISNUMBER
                FColumnHeaderFilterType = LvwColumnHeaderFilterTypeNumber
        End Select
        End With
    End If
End If
End Property

Friend Property Let FColumnHeaderFilterType(ByVal Index As Long, ByVal Value As LvwColumnHeaderFilterTypeConstants)
If ListViewHandle <> 0 Then
    ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Select Case Value
            Case LvwColumnHeaderFilterTypeText, LvwColumnHeaderFilterTypeNumber
                Dim HDI As HDITEM
                With HDI
                .Mask = HDI_FILTER
                .FilterType = Value
                Select Case .FilterType
                    Case HDFT_ISSTRING
                        Dim HDTF As HDTEXTFILTER
                        .pvFilter = VarPtr(HDTF)
                    Case HDFT_ISNUMBER
                        Dim LngValue As Long
                        .pvFilter = VarPtr(LngValue)
                End Select
                SendMessage ListViewHeaderHandle, HDM_SETITEM, Index - 1, ByVal VarPtr(HDI)
                .FilterType = .FilterType Or HDFT_HASNOVALUE
                .pvFilter = 0
                SendMessage ListViewHeaderHandle, HDM_SETITEM, Index - 1, ByVal VarPtr(HDI)
                End With
            Case Else
                Err.Raise 380
        End Select
    End If
End If
End Property

Friend Property Get FColumnHeaderFilterValue(ByVal Index As Long) As Variant
If ListViewHandle <> 0 Then
    ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim HDI As HDITEM
        With HDI
        .Mask = HDI_FILTER
        SendMessage ListViewHeaderHandle, HDM_GETITEM, Index - 1, ByVal VarPtr(HDI)
        If (.FilterType And HDFT_HASNOVALUE) = HDFT_HASNOVALUE Then
            FColumnHeaderFilterValue = Null
        Else
            Select Case .FilterType
                Case HDFT_ISSTRING
                    Dim HDTF As HDTEXTFILTER, Buffer As String
                    Buffer = String(MAX_PATH, vbNullChar) & vbNullChar
                    HDTF.pszText = StrPtr(Buffer)
                    HDTF.cchTextMax = Len(Buffer)
                    .pvFilter = VarPtr(HDTF)
                    SendMessage ListViewHeaderHandle, HDM_GETITEM, Index - 1, ByVal VarPtr(HDI)
                    FColumnHeaderFilterValue = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
                Case HDFT_ISNUMBER
                    Dim LngValue As Long
                    .pvFilter = VarPtr(LngValue)
                    SendMessage ListViewHeaderHandle, HDM_GETITEM, Index - 1, ByVal VarPtr(HDI)
                    FColumnHeaderFilterValue = LngValue
            End Select
        End If
        End With
    End If
End If
End Property

Friend Property Let FColumnHeaderFilterValue(ByVal Index As Long, ByVal Value As Variant)
If ListViewHandle <> 0 Then
    ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim HDI As HDITEM
        With HDI
        .Mask = HDI_FILTER
        SendMessage ListViewHeaderHandle, HDM_GETITEM, Index - 1, ByVal VarPtr(HDI)
        If (.FilterType And HDFT_HASNOVALUE) = HDFT_HASNOVALUE Then .FilterType = .FilterType And Not HDFT_HASNOVALUE
        Select Case .FilterType
            Case HDFT_ISSTRING
                Dim HDTF As HDTEXTFILTER
                Select Case VarType(Value)
                    Case vbString
                        HDTF.pszText = StrPtr(Value)
                        HDTF.cchTextMax = Len(Value)
                        .pvFilter = VarPtr(HDTF)
                    Case vbNull, vbEmpty
                        .FilterType = .FilterType Or HDFT_HASNOVALUE
                        .pvFilter = 0
                    Case Else
                        Err.Raise 13
                End Select
            Case HDFT_ISNUMBER
                Dim LngValue As Long
                Select Case VarType(Value)
                    Case vbLong, vbInteger, vbByte
                        LngValue = Value
                        .pvFilter = VarPtr(LngValue)
                    Case vbDouble, vbSingle
                        LngValue = CLng(Value)
                        .pvFilter = VarPtr(LngValue)
                    Case vbNull, vbEmpty
                        .FilterType = .FilterType Or HDFT_HASNOVALUE
                        .pvFilter = 0
                    Case Else
                        Err.Raise 13
                End Select
        End Select
        SendMessage ListViewHeaderHandle, HDM_SETITEM, Index - 1, ByVal VarPtr(HDI)
        End With
    End If
End If
End Property

Friend Property Get FColumnHeaderLeft(ByVal Index As Long) As Single
If ListViewHandle <> 0 Then
    Dim i As Long
    For i = 1 To Index
        If i = Index Then Exit For
        FColumnHeaderLeft = FColumnHeaderLeft + Me.FColumnHeaderWidth(i)
    Next i
End If
End Property

Friend Sub FColumnHeaderAutoSize(ByVal Index As Long, ByVal Value As LvwColumnHeaderAutoSizeConstants)
If ListViewHandle <> 0 Then
    Dim Flag As Long
    Select Case Value
        Case LvwColumnHeaderAutoSizeToItems
            Flag = LVSCW_AUTOSIZE
        Case LvwColumnHeaderAutoSizeToHeader
            Flag = LVSCW_AUTOSIZE_USEHEADER
        Case Else
            Err.Raise 380
    End Select
    SendMessage ListViewHandle, LVM_SETCOLUMNWIDTH, Index - 1, ByVal Flag
End If
End Sub

Friend Sub FColumnHeaderEditFilter(ByVal Index As Long)
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        UserControl.SetFocus
        SendMessage ListViewHeaderHandle, HDM_EDITFILTER, Index - 1, ByVal 0&
    End If
End If
End Sub

Friend Sub FColumnHeaderClearFilter(ByVal Index As Long)
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then SendMessage ListViewHeaderHandle, HDM_CLEARFILTER, Index - 1, ByVal 0&
End If
End Sub

Friend Function FColumnHeaderSubItemIndex(ByVal Index As Long) As Long
If ListViewHandle <> 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_SUBITEM
    SendMessage ListViewHandle, LVM_GETCOLUMN, Index - 1, ByVal VarPtr(LVC)
    FColumnHeaderSubItemIndex = LVC.iSubItem
End If
End Function

Public Property Get Groups() As LvwGroups
Attribute Groups.VB_Description = "Returns a reference to a collection of the group objects. Any groups assigned appear whenever the view property is other than 'list' view. Requires comctl32.dll version 6.0 or higher."
If PropGroups Is Nothing Then
    If ComCtlsSupportLevel() >= 1 Then
        If PropVirtualMode = False Then
            Set PropGroups = New LvwGroups
            PropGroups.FInit Me
        Else
            Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
        End If
    Else
        Err.Raise Number:=91, Description:="To use this functionality, you must provide a manifest specifying comctl32.dll version 6.0 or higher."
    End If
End If
Set Groups = PropGroups
End Property

Friend Sub FGroupsAdd(ByVal Index As Long, ByVal NewGroup As LvwGroup, ByVal This As ISubclass, Optional ByVal Header As String, Optional ByVal HeaderAlignment As LvwGroupHeaderAlignmentConstants, Optional ByVal Footer As String, Optional ByVal FooterAlignment As LvwGroupFooterAlignmentConstants)
If ComCtlsSupportLevel() = 0 Then Exit Sub
Dim LVG As LVGROUP
With LVG
.cbSize = LenB(LVG)
.Mask = LVGF_GROUPID Or LVGF_ALIGN
.iGroupId = NextGroupID()
NewGroup.ID = .iGroupId
If Not Header = vbNullString Then
    .Mask = .Mask Or LVGF_HEADER
    .pszHeader = StrPtr(Header)
    .cchHeader = Len(Header) + 1
End If
Select Case HeaderAlignment
    Case LvwGroupHeaderAlignmentLeft
        .uAlign = LVGA_HEADER_LEFT
    Case LvwGroupHeaderAlignmentRight
        .uAlign = LVGA_HEADER_RIGHT
    Case LvwGroupHeaderAlignmentCenter
        .uAlign = LVGA_HEADER_CENTER
End Select
If ComCtlsSupportLevel() >= 2 Then
    If Not Footer = vbNullString Then
        .Mask = .Mask Or LVGF_FOOTER
        .pszFooter = StrPtr(Footer)
        .cchFooter = Len(Footer) + 1
    End If
    Select Case FooterAlignment
        Case LvwGroupFooterAlignmentLeft
            .uAlign = .uAlign Or LVGA_FOOTER_LEFT
        Case LvwGroupFooterAlignmentRight
            .uAlign = .uAlign Or LVGA_FOOTER_RIGHT
        Case LvwGroupFooterAlignmentCenter
            .uAlign = .uAlign Or LVGA_FOOTER_CENTER
    End Select
End If
End With
If ListViewHandle <> 0 Then
    If This Is Nothing Then
        SendMessage ListViewHandle, LVM_INSERTGROUP, Index - 1, ByVal VarPtr(LVG)
    Else
        Dim LVIGS As LVINSERTGROUPSORTED
        With LVIGS
        LSet .LVG = LVG
        Set .pvData = This
        .pfnGroupCompare = ProcPtr(AddressOf ComCtlsLvwSortingFunctionGroups)
        End With
        SendMessage ListViewHandle, LVM_INSERTGROUPSORTED, VarPtr(LVIGS), ByVal 0&
    End If
End If
End Sub

Friend Sub FGroupsRemove(ByVal ID As Long)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then SendMessage ListViewHandle, LVM_REMOVEGROUP, ID, ByVal 0&
End Sub

Friend Sub FGroupsClear()
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    SendMessage ListViewHandle, LVM_REMOVEALLGROUPS, 0, ByVal 0&
    ' LVM_REMOVEALLGROUPS turns off the group view.
    ' Thus it is necessary to reapply the group view property.
    Me.GroupView = PropGroupView
End If
End Sub

Friend Sub FGroupsSort(ByVal This As ISubclass)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then SendMessage ListViewHandle, LVM_SORTGROUPS, ProcPtr(AddressOf ComCtlsLvwSortingFunctionGroups), ByVal ObjPtr(This)
End Sub

Friend Property Get FGroupHeader(ByVal ID As Long) As String
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_HEADER
        Dim Buffer As String
        Buffer = String(260, vbNullChar)
        .pszHeader = StrPtr(Buffer)
        .cchHeader = 260
        End With
        SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG)
        FGroupHeader = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End If
End If
End Property

Friend Property Let FGroupHeader(ByVal ID As Long, ByVal Value As String)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_HEADER
        .pszHeader = StrPtr(Value)
        .cchHeader = Len(Value) + 1
        End With
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
    End If
End If
End Property

Friend Property Get FGroupHeaderAlignment(ByVal ID As Long) As LvwGroupHeaderAlignmentConstants
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_ALIGN
        SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG)
        If (.uAlign And LVGA_HEADER_LEFT) = LVGA_HEADER_LEFT Then
            FGroupHeaderAlignment = LvwGroupHeaderAlignmentLeft
        ElseIf (.uAlign And LVGA_HEADER_RIGHT) = LVGA_HEADER_RIGHT Then
            FGroupHeaderAlignment = LvwGroupHeaderAlignmentRight
        ElseIf (.uAlign And LVGA_HEADER_CENTER) = LVGA_HEADER_CENTER Then
            FGroupHeaderAlignment = LvwGroupHeaderAlignmentCenter
        End If
        End With
    End If
End If
End Property

Friend Property Let FGroupHeaderAlignment(ByVal ID As Long, ByVal Value As LvwGroupHeaderAlignmentConstants)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If IsGroupAvailable(ID) = True Then
        Select Case Value
            Case LvwGroupHeaderAlignmentLeft, LvwGroupHeaderAlignmentRight, LvwGroupHeaderAlignmentCenter
                Dim LVG As LVGROUP
                With LVG
                .cbSize = LenB(LVG)
                .Mask = LVGF_ALIGN
                SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG)
                If (.uAlign And LVGA_HEADER_LEFT) = LVGA_HEADER_LEFT Then .uAlign = .uAlign And Not LVGA_HEADER_LEFT
                If (.uAlign And LVGA_HEADER_RIGHT) = LVGA_HEADER_RIGHT Then .uAlign = .uAlign And Not LVGA_HEADER_RIGHT
                If (.uAlign And LVGA_HEADER_CENTER) = LVGA_HEADER_CENTER Then .uAlign = .uAlign And Not LVGA_HEADER_CENTER
                Select Case Value
                    Case LvwGroupHeaderAlignmentLeft
                        .uAlign = .uAlign Or LVGA_HEADER_LEFT
                    Case LvwGroupHeaderAlignmentRight
                        .uAlign = .uAlign Or LVGA_HEADER_RIGHT
                    Case LvwGroupHeaderAlignmentCenter
                        .uAlign = .uAlign Or LVGA_HEADER_CENTER
                End Select
                End With
                SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
                SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
                If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
            Case Else
                Err.Raise 380
        End Select
    End If
End If
End Property

Friend Property Get FGroupFooter(ByVal ID As Long) As String
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_FOOTER
        Dim Buffer As String
        Buffer = String(260, vbNullChar)
        .pszFooter = StrPtr(Buffer)
        .cchFooter = 260
        End With
        SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG)
        FGroupFooter = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End If
End If
End Property

Friend Property Let FGroupFooter(ByVal ID As Long, ByVal Value As String)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_FOOTER
        .pszFooter = StrPtr(Value)
        .cchFooter = Len(Value) + 1
        End With
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
    End If
End If
End Property

Friend Property Get FGroupFooterAlignment(ByVal ID As Long) As LvwGroupFooterAlignmentConstants
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_ALIGN
        SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG)
        If (.uAlign And LVGA_FOOTER_LEFT) = LVGA_FOOTER_LEFT Then
            FGroupFooterAlignment = LvwGroupFooterAlignmentLeft
        ElseIf (.uAlign And LVGA_FOOTER_RIGHT) = LVGA_FOOTER_RIGHT Then
            FGroupFooterAlignment = LvwGroupFooterAlignmentRight
        ElseIf (.uAlign And LVGA_FOOTER_CENTER) = LVGA_FOOTER_CENTER Then
            FGroupFooterAlignment = LvwGroupFooterAlignmentCenter
        End If
        End With
    End If
End If
End Property

Friend Property Let FGroupFooterAlignment(ByVal ID As Long, ByVal Value As LvwGroupFooterAlignmentConstants)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Select Case Value
            Case LvwGroupFooterAlignmentLeft, LvwGroupFooterAlignmentRight, LvwGroupFooterAlignmentCenter
                Dim LVG As LVGROUP
                With LVG
                .cbSize = LenB(LVG)
                .Mask = LVGF_ALIGN
                SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG)
                If (.uAlign And LVGA_FOOTER_LEFT) = LVGA_FOOTER_LEFT Then .uAlign = .uAlign And Not LVGA_FOOTER_LEFT
                If (.uAlign And LVGA_FOOTER_RIGHT) = LVGA_FOOTER_RIGHT Then .uAlign = .uAlign And Not LVGA_FOOTER_RIGHT
                If (.uAlign And LVGA_FOOTER_CENTER) = LVGA_FOOTER_CENTER Then .uAlign = .uAlign And Not LVGA_FOOTER_CENTER
                Select Case Value
                    Case LvwGroupFooterAlignmentLeft
                        .uAlign = .uAlign Or LVGA_FOOTER_LEFT
                    Case LvwGroupFooterAlignmentRight
                        .uAlign = .uAlign Or LVGA_FOOTER_RIGHT
                    Case LvwGroupFooterAlignmentCenter
                        .uAlign = .uAlign Or LVGA_FOOTER_CENTER
                End Select
                End With
                SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
                SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
                If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
            Case Else
                Err.Raise 380
        End Select
    End If
End If
End Property

Friend Property Get FGroupHint(ByVal ID As Long) As String
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG_V61 As LVGROUP_V61
        With LVG_V61
        .LVG.cbSize = LenB(LVG_V61)
        .LVG.Mask = LVGF_SUBTITLE
        Dim Buffer As String
        Buffer = String(260, vbNullChar)
        .pszSubtitle = StrPtr(Buffer)
        .cchSubtitle = 260
        End With
        SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG_V61)
        FGroupHint = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End If
End If
End Property

Friend Property Let FGroupHint(ByVal ID As Long, ByVal Value As String)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG_V61 As LVGROUP_V61
        With LVG_V61
        .LVG.cbSize = LenB(LVG_V61)
        .LVG.Mask = LVGF_SUBTITLE
        .pszSubtitle = StrPtr(Value)
        .cchSubtitle = Len(Value) + 1
        End With
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG_V61)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
    End If
End If
End Property

Friend Property Get FGroupLink(ByVal ID As Long) As String
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG_V61 As LVGROUP_V61
        With LVG_V61
        .LVG.cbSize = LenB(LVG_V61)
        .LVG.Mask = LVGF_TASK
        Dim Buffer As String
        Buffer = String(260, vbNullChar)
        .pszTask = StrPtr(Buffer)
        .cchTask = 260
        End With
        SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG_V61)
        FGroupLink = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End If
End If
End Property

Friend Property Let FGroupLink(ByVal ID As Long, ByVal Value As String)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG_V61 As LVGROUP_V61
        With LVG_V61
        .LVG.cbSize = LenB(LVG_V61)
        .LVG.Mask = LVGF_TASK
        .pszTask = StrPtr(Value)
        .cchTask = Len(Value) + 1
        End With
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG_V61)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
    End If
End If
End Property

Friend Property Get FGroupSubsetLink(ByVal ID As Long) As String
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG_V61 As LVGROUP_V61
        With LVG_V61
        .LVG.cbSize = LenB(LVG_V61)
        .LVG.Mask = LVGF_SUBSET Or LVGF_SUBSETITEMS
        Dim Buffer As String
        Buffer = String(260, vbNullChar)
        .pszSubsetTitle = StrPtr(Buffer)
        .cchSubsetTitle = 260
        End With
        SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG_V61)
        FGroupSubsetLink = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
    End If
End If
End Property

Friend Property Let FGroupSubsetLink(ByVal ID As Long, ByVal Value As String)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG_V61 As LVGROUP_V61
        With LVG_V61
        .LVG.cbSize = LenB(LVG_V61)
        .LVG.Mask = LVGF_SUBSET Or LVGF_SUBSETITEMS
        .pszSubsetTitle = StrPtr(Value)
        .cchSubsetTitle = Len(Value) + 1
        End With
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG_V61)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
    End If
End If
End Property

Friend Property Get FGroupCollapsible(ByVal ID As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then FGroupCollapsible = CBool(SendMessage(ListViewHandle, LVM_GETGROUPSTATE, ID, ByVal LVGS_COLLAPSIBLE) <> 0)
End If
End Property

Friend Property Let FGroupCollapsible(ByVal ID As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_STATE
        .StateMask = LVGS_COLLAPSIBLE
        If Value = True Then
            .State = LVGS_COLLAPSIBLE
        Else
            .State = LVGS_NORMAL
        End If
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
        End With
    End If
End If
End Property

Friend Property Get FGroupCollapsed(ByVal ID As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then FGroupCollapsed = CBool(SendMessage(ListViewHandle, LVM_GETGROUPSTATE, ID, ByVal LVGS_COLLAPSED) <> 0)
End If
End Property

Friend Property Let FGroupCollapsed(ByVal ID As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_STATE
        .StateMask = LVGS_COLLAPSED
        If Value = True Then
            .State = LVGS_COLLAPSED
        Else
            .State = LVGS_NORMAL
        End If
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
        End With
    End If
End If
End Property

Friend Property Get FGroupShowHeader(ByVal ID As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then FGroupShowHeader = CBool(SendMessage(ListViewHandle, LVM_GETGROUPSTATE, ID, ByVal LVGS_NOHEADER) = 0)
End If
End Property

Friend Property Let FGroupShowHeader(ByVal ID As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_STATE
        .StateMask = LVGS_NOHEADER
        If Value = True Then
            .State = LVGS_NORMAL
        Else
            .State = LVGS_NOHEADER
        End If
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
        End With
    End If
End If
End Property

Friend Property Get FGroupSelected(ByVal ID As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then FGroupSelected = CBool(SendMessage(ListViewHandle, LVM_GETGROUPSTATE, ID, ByVal LVGS_SELECTED) <> 0)
End If
End Property

Friend Property Let FGroupSelected(ByVal ID As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_STATE
        If Value = True Then
            .StateMask = LVGS_SELECTED Or LVGS_FOCUSED
            .State = LVGS_SELECTED Or LVGS_FOCUSED
        Else
            .StateMask = LVGS_SELECTED
            .State = LVGS_NORMAL
        End If
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
        End With
    End If
End If
End Property

Friend Property Get FGroupSubseted(ByVal ID As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then FGroupSubseted = CBool(SendMessage(ListViewHandle, LVM_GETGROUPSTATE, ID, ByVal LVGS_SUBSETED) <> 0)
End If
End Property

Friend Property Let FGroupSubseted(ByVal ID As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_STATE
        If Value = True Then
            .StateMask = LVGS_SUBSETED
            .State = LVGS_SUBSETED
        Else
            .StateMask = LVGS_SUBSETED
            .State = LVGS_NORMAL
        End If
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
        End With
    End If
End If
End Property

Friend Property Get FGroupSubsetLinkSelected(ByVal ID As Long) As Boolean
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then FGroupSubsetLinkSelected = CBool(SendMessage(ListViewHandle, LVM_GETGROUPSTATE, ID, ByVal LVGS_SUBSETLINKFOCUSED) <> 0)
End If
End Property

Friend Property Let FGroupSubsetLinkSelected(ByVal ID As Long, ByVal Value As Boolean)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_STATE
        If Value = True Then
            .StateMask = LVGS_SUBSETLINKFOCUSED
            .State = LVGS_SUBSETLINKFOCUSED
        Else
            .StateMask = LVGS_SUBSETLINKFOCUSED
            .State = LVGS_NORMAL
        End If
        SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG)
        If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
        End With
    End If
End If
End Property

Friend Property Get FGroupIcon(ByVal ID As Long) As Long
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG_V61 As LVGROUP_V61
        With LVG_V61
        .LVG.cbSize = LenB(LVG_V61)
        .LVG.Mask = LVGF_TITLEIMAGE
        SendMessage ListViewHandle, LVM_GETGROUPINFO, ID, ByVal VarPtr(LVG_V61)
        FGroupIcon = .iTitleImage + 1
        End With
    End If
End If
End Property

Friend Property Let FGroupIcon(ByVal ID As Long, ByVal Value As Long)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim LVG_V61 As LVGROUP_V61
        With LVG_V61
        .LVG.cbSize = LenB(LVG_V61)
        .LVG.Mask = LVGF_TITLEIMAGE
        .iTitleImage = Value - 1
        SendMessage ListViewHandle, LVM_SETGROUPINFO, ID, ByVal VarPtr(LVG_V61)
        End With
    End If
End If
End Property

Friend Property Get FGroupPosition(ByVal ID As Long) As Long
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim Count As Long
        Count = SendMessage(ListViewHandle, LVM_GETGROUPCOUNT, 0, ByVal 0&)
        If Count > 0 Then
            Dim LVG As LVGROUP, i As Long
            With LVG
            .cbSize = LenB(LVG)
            .Mask = LVGF_GROUPID
            For i = 0 To Count - 1
                SendMessage ListViewHandle, LVM_GETGROUPINFOBYINDEX, i, ByVal VarPtr(LVG)
                If .iGroupId = ID Then
                    FGroupPosition = i + 1
                    Exit For
                End If
            Next i
            End With
        End If
    End If
End If
End Property

Friend Property Get FGroupLeft(ByVal ID As Long) As Single
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim RC As RECT
        RC.Top = LVGGR_HEADER
        SendMessage ListViewHandle, LVM_GETGROUPRECT, ID, ByVal VarPtr(RC)
        FGroupLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
    End If
End If
End Property

Friend Property Get FGroupTop(ByVal ID As Long) As Single
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim RC As RECT
        RC.Top = LVGGR_HEADER
        SendMessage ListViewHandle, LVM_GETGROUPRECT, ID, ByVal VarPtr(RC)
        FGroupTop = UserControl.ScaleY(RC.Top, vbPixels, vbContainerPosition)
    End If
End If
End Property

Friend Property Get FGroupWidth(ByVal ID As Long) As Single
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim RC As RECT
        RC.Top = LVGGR_HEADER
        SendMessage ListViewHandle, LVM_GETGROUPRECT, ID, ByVal VarPtr(RC)
        FGroupWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
    End If
End If
End Property

Friend Property Get FGroupHeight(ByVal ID As Long) As Single
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If IsGroupAvailable(ID) = True Then
        Dim RC As RECT
        RC.Top = LVGGR_HEADER
        SendMessage ListViewHandle, LVM_GETGROUPRECT, ID, ByVal VarPtr(RC)
        FGroupHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
    End If
End If
End Property

Friend Property Get FGroupListItemIndices(ByVal ID As Long) As Collection
Set FGroupListItemIndices = New Collection
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If IsGroupAvailable(ID) = True Then
        If SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&) > 0 Then
            Dim LVI_V60 As LVITEM_V60
            With LVI_V60
            .LVI.Mask = LVIF_GROUPID
            .LVI.iItem = 0
            SendMessage ListViewHandle, LVM_GETITEM, 0, ByVal VarPtr(LVI_V60)
            Do While .LVI.iItem > -1
                If .iGroupId = ID Then FGroupListItemIndices.Add (.LVI.iItem + 1)
                .LVI.iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, .LVI.iItem, ByVal LVNI_ALL)
                SendMessage ListViewHandle, LVM_GETITEM, 0, ByVal VarPtr(LVI_V60)
            Loop
            End With
        End If
    End If
End If
End Property

Public Property Get WorkAreas() As LvwWorkAreas
Attribute WorkAreas.VB_Description = "Returns a reference to a collection of the work area objects."
If PropWorkAreas Is Nothing Then
    If PropVirtualMode = False Then
        Set PropWorkAreas = New LvwWorkAreas
        PropWorkAreas.FInit Me
    Else
        Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
    End If
End If
Set WorkAreas = PropWorkAreas
End Property

Friend Function FWorkAreasAdd(ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, Optional ByVal Index As Long) As Long
Dim Count As Long, WorkAreaIndex As Long
Count = Me.FWorkAreasCount
If Index = 0 Then
    WorkAreaIndex = Count + 1
Else
    WorkAreaIndex = Index
End If
Dim RC As RECT
RC.Left = UserControl.ScaleX(Left, vbContainerPosition, vbPixels)
RC.Top = UserControl.ScaleY(Top, vbContainerPosition, vbPixels)
RC.Right = RC.Left + UserControl.ScaleX(Width, vbContainerSize, vbPixels)
RC.Bottom = RC.Top + UserControl.ScaleY(Height, vbContainerSize, vbPixels)
If (RC.Right - RC.Left) > 0 And (RC.Bottom - RC.Top) > 0 Then
    If ListViewHandle <> 0 Then
        If Count < LV_MAX_WORKAREAS Then
            Dim ArrRC() As RECT, i As Long
            ReDim ArrRC(1 To (Count + 1)) As RECT
            SendMessage ListViewHandle, LVM_GETWORKAREAS, Count, ByVal VarPtr(ArrRC(1))
            Count = Count + 1
            If WorkAreaIndex < Count Then
                For i = Count To (WorkAreaIndex + 1) Step -1
                    LSet ArrRC(i) = ArrRC(i - 1)
                Next i
            End If
            LSet ArrRC(WorkAreaIndex) = RC
            SendMessage ListViewHandle, LVM_SETWORKAREAS, Count, ByVal VarPtr(ArrRC(1))
            FWorkAreasAdd = WorkAreaIndex
        Else
            ' The maximum number of work areas was exceeded. (Index out of bounds)
            FWorkAreasAdd = 0
        End If
    End If
Else
    ' Zero width or height is not accepted by LVM_SETWORKAREAS. (Invalid property value)
    FWorkAreasAdd = -1
End If
End Function

Friend Function FWorkAreasCount() As Long
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_GETNUMBEROFWORKAREAS, 0, ByVal VarPtr(FWorkAreasCount)
End Function

Friend Sub FWorkAreasClear()
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_SETWORKAREAS, 0, ByVal 0&
End Sub

Friend Sub FWorkAreasRemove(ByVal Index As Long)
Dim Count As Long
Count = Me.FWorkAreasCount
If Count > 0 And Index <= Count And Index > 0 Then
    If ListViewHandle <> 0 Then
        Dim ArrRC() As RECT
        ReDim ArrRC(1 To Count) As RECT
        SendMessage ListViewHandle, LVM_GETWORKAREAS, Count, ByVal VarPtr(ArrRC(1))
        Dim i As Long
        For i = Index To (Count - 1)
            LSet ArrRC(i) = ArrRC(i + 1)
        Next i
        Count = Count - 1
        If Count > 0 Then
            SendMessage ListViewHandle, LVM_SETWORKAREAS, Count, ByVal VarPtr(ArrRC(1))
        Else
            SendMessage ListViewHandle, LVM_SETWORKAREAS, 0, ByVal 0&
        End If
    End If
End If
End Sub

Friend Property Get FWorkAreaLeft(ByVal Index As Long) As Single
Dim RC As RECT
Call GetWorkAreaRect(Index, RC)
FWorkAreaLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerPosition)
End Property

Friend Property Get FWorkAreaTop(ByVal Index As Long) As Single
Dim RC As RECT
Call GetWorkAreaRect(Index, RC)
FWorkAreaTop = UserControl.ScaleY(RC.Top, vbPixels, vbContainerPosition)
End Property

Friend Property Get FWorkAreaWidth(ByVal Index As Long) As Single
Dim RC As RECT
Call GetWorkAreaRect(Index, RC)
FWorkAreaWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End Property

Friend Property Get FWorkAreaHeight(ByVal Index As Long) As Single
Dim RC As RECT
Call GetWorkAreaRect(Index, RC)
FWorkAreaHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End Property

Friend Property Get FWorkAreaListItemIndices(ByVal Index As Long) As Collection
Select Case PropView
    Case LvwViewIcon, LvwViewSmallIcon, LvwViewTile
        Set FWorkAreaListItemIndices = New Collection
        If ListViewHandle <> 0 Then
            Dim Count As Long
            SendMessage ListViewHandle, LVM_GETNUMBEROFWORKAREAS, 0, ByVal VarPtr(Count)
            If Count > 0 And Index <= Count And Index > 0 Then
                Dim ArrRC() As RECT
                ReDim ArrRC(1 To Count) As RECT
                SendMessage ListViewHandle, LVM_GETWORKAREAS, Count, ByVal VarPtr(ArrRC(1))
                Dim iItem As Long, P As POINTAPI, iWorkArea As Long
                For iItem = 0 To (SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&) - 1)
                    If SendMessage(ListViewHandle, LVM_GETITEMPOSITION, iItem, ByVal VarPtr(P)) <> 0 Then
                        For iWorkArea = 1 To Index
                            If PtInRect(ArrRC(iWorkArea), P.X, P.Y) <> 0 Then Exit For
                        Next iWorkArea
                        If iWorkArea = Index Then FWorkAreaListItemIndices.Add (iItem + 1)
                    End If
                Next iItem
            End If
        End If
    Case Else
        Err.Raise Number:=394, Description:="Get supported in 'icon', 'small icon' and 'tile' view only"
End Select
End Property

Private Sub CreateListView()
If ListViewHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or LVS_SHAREIMAGELISTS
If PropRightToLeft = True Then
    If PropRightToLeftLayout = True Then
        dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    Else
        dwExStyle = dwExStyle Or WS_EX_RTLREADING
    End If
End If
Call ComCtlsInitBorderStyle(dwStyle, dwExStyle, PropBorderStyle)
If ListViewDesignMode = False Then
    If (ComCtlsSupportLevel() = 0 Or PropVirtualMode = True) And PropView = LvwViewTile Then PropView = LvwViewIcon
    Select Case PropView
        Case LvwViewIcon
            dwStyle = dwStyle Or LVS_ICON
        Case LvwViewSmallIcon
            dwStyle = dwStyle Or LVS_SMALLICON
        Case LvwViewList
            dwStyle = dwStyle Or LVS_LIST
        Case LvwViewReport
            dwStyle = dwStyle Or LVS_REPORT
    End Select
    If PropVirtualMode = False Then
        Select Case PropArrange
            Case LvwArrangeAutoLeft
                dwStyle = dwStyle Or LVS_AUTOARRANGE Or LVS_ALIGNLEFT
            Case LvwArrangeAutoTop
                dwStyle = dwStyle Or LVS_AUTOARRANGE Or LVS_ALIGNTOP
            Case LvwArrangeLeft
                dwStyle = dwStyle Or LVS_ALIGNLEFT
            Case LvwArrangeTop
                dwStyle = dwStyle Or LVS_ALIGNTOP
        End Select
    Else
        ' According to MSDN:
        ' All virtual list view controls default to the LVS_AUTOARRANGE style.
        dwStyle = dwStyle Or LVS_AUTOARRANGE
    End If
Else
    dwStyle = dwStyle Or LVS_LIST
End If
If PropMultiSelect = False Then dwStyle = dwStyle Or LVS_SINGLESEL
If PropLabelEdit <> LvwLabelEditDisabled Then dwStyle = dwStyle Or LVS_EDITLABELS
If PropLabelWrap = False Then dwStyle = dwStyle Or LVS_NOLABELWRAP
If PropHideSelection = False Then dwStyle = dwStyle Or LVS_SHOWSELALWAYS
If PropHideColumnHeaders = True Then dwStyle = dwStyle Or LVS_NOCOLUMNHEADER
If ListViewDesignMode = False Then
    If PropVirtualMode = True Then dwStyle = dwStyle Or LVS_OWNERDATA
    ' The WM_NOTIFYFORMAT notification must be handled, which will be sent on control creation.
    ' Thus it is necessary to subclass the parent before the control is created.
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 5)
End If
ListViewHandle = CreateWindowEx(dwExStyle, StrPtr("SysListView32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If ListViewHandle <> 0 Then
    ListViewToolTipHandle = SendMessage(ListViewHandle, LVM_GETTOOLTIPS, 0, ByVal 0&)
    If ListViewToolTipHandle <> 0 Then Call ComCtlsInitToolTip(ListViewToolTipHandle)
End If
If PropView = LvwViewReport Then
    ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then Call ComCtlsSetSubclass(ListViewHeaderHandle, Me, 4)
End If
If ListViewHandle <> 0 Then
    If ListViewDesignMode = False Then
        If PropView = LvwViewTile Then SendMessage ListViewHandle, LVM_SETVIEW, LV_VIEW_TILE, ByVal 0&
        SendMessage ListViewHandle, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_SUBITEMIMAGES, ByVal LVS_EX_SUBITEMIMAGES
        If PropVirtualMode = True Then
            Dim CallbackMask As Long
            CallbackMask = SendMessage(ListViewHandle, LVM_GETCALLBACKMASK, 0, ByVal 0&)
            If (CallbackMask And LVIS_STATEIMAGEMASK) = 0 Then CallbackMask = CallbackMask Or LVIS_STATEIMAGEMASK
            SendMessage ListViewHandle, LVM_SETCALLBACKMASK, CallbackMask, ByVal 0&
        End If
    End If
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.HotMousePointer = PropHotMousePointer
Me.BackColor = PropBackColor
Me.ForeColor = PropForeColor
If PropRedraw = False Then Me.Redraw = False
Me.AllowColumnReorder = PropAllowColumnReorder
Me.AllowColumnCheckboxes = PropAllowColumnCheckboxes
Me.FullRowSelect = PropFullRowSelect
Me.GridLines = PropGridLines
Me.Checkboxes = PropCheckboxes
Me.ShowInfoTips = PropShowInfoTips
Me.ShowLabelTips = PropShowLabelTips
Me.ShowColumnTips = PropShowColumnTips
Me.DoubleBuffer = PropDoubleBuffer
Me.HoverSelection = PropHoverSelection
Me.HoverSelectionTime = PropHoverSelectionTime
Me.HotTracking = PropHotTracking
Me.InsertMarkColor = PropInsertMarkColor
Me.TextBackground = PropTextBackground
Me.ClickableColumnHeaders = PropClickableColumnHeaders
Me.HighlightColumnHeaders = PropHighlightColumnHeaders
Me.TrackSizeColumnHeaders = PropTrackSizeColumnHeaders
Me.ResizableColumnHeaders = PropResizableColumnHeaders
If Not PropPicture Is Nothing Then Set Me.Picture = PropPicture
Me.TileViewLines = PropTileViewLines
Me.SnapToGrid = PropSnapToGrid
Me.GroupView = PropGroupView
Me.GroupSubsetCount = PropGroupSubsetCount
Me.UseColumnChevron = PropUseColumnChevron
Me.UseColumnFilterBar = PropUseColumnFilterBar
Me.VirtualItemCount = PropVirtualItemCount
If ListViewHandle <> 0 Then
    If ComCtlsSupportLevel() = 0 Then
        ' According to MSDN:
        ' - Version 5 of comctl32 supports deleting of column zero, but only after you use CCM_SETVERSION to set the version to 5 or later.
        ' - If you change the font by returning CDRF_NEWFONT, the list view control might display clipped text.
        '   This behavior is necessary for backward compatibility with earlier versions of the common controls.
        '   If you want to change the font of a list view control, you will get better results if you send a CCM_SETVERSION message
        '   with the wParam value set to 5 before adding any items to the control.
        SendMessage ListViewHandle, CCM_SETVERSION, 5, ByVal 0&
    End If
End If
If ListViewDesignMode = False Then
    If ListViewHandle <> 0 Then
        Call ComCtlsSetSubclass(ListViewHandle, Me, 1)
        Call ComCtlsCreateIMC(ListViewHandle, ListViewIMCHandle)
    End If
End If
End Sub

Private Sub CreateHeaderToolTip()
Static Done As Boolean
If ListViewHeaderToolTipHandle <> 0 Then Exit Sub
If Done = False Then
    Call ComCtlsInitCC(ICC_TAB_CLASSES)
    Done = True
End If
Dim dwExStyle As Long
dwExStyle = WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
ListViewHeaderToolTipHandle = CreateWindowEx(dwExStyle, StrPtr("tooltips_class32"), StrPtr("Tool Tip"), WS_POPUP Or TTS_ALWAYSTIP Or TTS_NOPREFIX, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If ListViewHeaderToolTipHandle <> 0 Then
    Call ComCtlsInitToolTip(ListViewHeaderToolTipHandle)
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = ListViewHeaderHandle
    .uId = 0
    .uFlags = TTF_SUBCLASS Or TTF_PARSELINKS
    If PropRightToLeft = True And PropRightToLeftLayout = False Then .uFlags = .uFlags Or TTF_RTLREADING
    .lpszText = LPSTR_TEXTCALLBACK
    GetClientRect ListViewHeaderHandle, .RC
    End With
    SendMessage ListViewHeaderToolTipHandle, TTM_ADDTOOL, 0, ByVal VarPtr(TI)
End If
Call SetVisualStylesHeaderToolTip
End Sub

Private Sub DestroyListView()
If ListViewHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(ListViewHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
Call ComCtlsDestroyIMC(ListViewHandle, ListViewIMCHandle)
Call DestroyHeaderToolTip
ShowWindow ListViewHandle, SW_HIDE
SetParent ListViewHandle, 0
DestroyWindow ListViewHandle
ListViewHandle = 0
ListViewHeaderHandle = 0
ListViewToolTipHandle = 0
If ListViewFontHandle <> 0 Then
    DeleteObject ListViewFontHandle
    ListViewFontHandle = 0
End If
If ListViewBoldFontHandle <> 0 Then
    DeleteObject ListViewBoldFontHandle
    ListViewBoldFontHandle = 0
End If
If ListViewUnderlineFontHandle <> 0 Then
    DeleteObject ListViewUnderlineFontHandle
    ListViewUnderlineFontHandle = 0
End If
If ListViewBoldUnderlineFontHandle <> 0 Then
    DeleteObject ListViewBoldUnderlineFontHandle
    ListViewBoldUnderlineFontHandle = 0
End If
End Sub

Private Sub DestroyHeaderToolTip()
If ListViewHeaderToolTipHandle = 0 Then Exit Sub
DestroyWindow ListViewHeaderToolTipHandle
ListViewHeaderToolTipHandle = 0
ListViewHeaderToolTipItem = -1
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
If PropRedraw = True Or ListViewDesignMode = True Then RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Function HitTest(ByVal X As Single, ByVal Y As Single, Optional ByRef SubItemIndex As Variant) As LvwListItem
Attribute HitTest.VB_Description = "Returns a reference to the list item object located at the coordinates of X and Y."
If ListViewHandle <> 0 Then
    Dim LVHTI As LVHITTESTINFO
    With LVHTI
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    If IsMissing(SubItemIndex) = True Then
        If SendMessage(ListViewHandle, LVM_HITTEST, 0, ByVal VarPtr(LVHTI)) > -1 Then
            If (.Flags And LVHT_ONITEM) <> 0 Then
                If PropVirtualMode = False Then
                    Set HitTest = Me.ListItems(.iItem + 1)
                Else
                    Set HitTest = New LvwListItem
                    HitTest.FInit ObjPtr(Me), .iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                End If
            End If
        End If
    Else
        If SendMessage(ListViewHandle, LVM_SUBITEMHITTEST, 0, ByVal VarPtr(LVHTI)) > -1 Then
            If (.Flags And LVHT_ONITEM) <> 0 Then
                If PropVirtualMode = False Then
                    Set HitTest = Me.ListItems(.iItem + 1)
                Else
                    Set HitTest = New LvwListItem
                    HitTest.FInit ObjPtr(Me), .iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                End If
                SubItemIndex = .iSubItem
            End If
        End If
    End If
    End With
End If
End Function

Public Function HitTestInsertMark(ByVal X As Single, ByVal Y As Single, Optional ByRef After As Boolean) As LvwListItem
Attribute HitTestInsertMark.VB_Description = "Returns a reference to the list item object located at the coordinates of X and Y and retrieves a value that determines where the insertion point should appear. Requires comctl32.dll version 6.1 or higher."
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim P As POINTAPI, LVIM As LVINSERTMARK
    P.X = CLng(UserControl.ScaleX(X, vbContainerPosition, vbPixels))
    P.Y = CLng(UserControl.ScaleY(Y, vbContainerPosition, vbPixels))
    With LVIM
    .cbSize = LenB(LVIM)
    SendMessage ListViewHandle, LVM_INSERTMARKHITTEST, VarPtr(P), ByVal VarPtr(LVIM)
    If .iItem > -1 Then
        If PropVirtualMode = False Then
            Set HitTestInsertMark = Me.ListItems(.iItem + 1)
        Else
            Set HitTestInsertMark = New LvwListItem
            HitTestInsertMark.FInit ObjPtr(Me), .iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
        End If
    End If
    After = CBool((.dwFlags And LVIM_AFTER) <> 0)
    End With
End If
End Function

Public Function FindItem(ByVal Text As String, Optional ByVal Index As Long, Optional ByVal Partial As Boolean, Optional ByVal Wrap As Boolean) As LvwListItem
Attribute FindItem.VB_Description = "Finds an item in the list and returns a reference to that item."
If ListViewHandle <> 0 Then
    If Index >= 0 Then
        Dim Count As Long
        Count = SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&)
        If Count > 0 Then
            If Index <= Count Then
                If Index > 0 Then Index = Index - 1
                Dim LVFI As LVFINDINFO
                With LVFI
                .psz = StrPtr(Text)
                .Flags = LVFI_STRING
                If Partial = True Then .Flags = .Flags Or LVFI_PARTIAL
                If Wrap = True Then .Flags = .Flags Or LVFI_WRAP
                End With
                Index = SendMessage(ListViewHandle, LVM_FINDITEM, Index - 1, ByVal VarPtr(LVFI))
                If Index > -1 Then
                    If PropVirtualMode = False Then
                        Set FindItem = Me.ListItems(Index + 1)
                    Else
                        Set FindItem = New LvwListItem
                        FindItem.FInit ObjPtr(Me), Index + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                    End If
                End If
            Else
                Err.Raise 380
            End If
        End If
    Else
        Err.Raise 380
    End If
End If
End Function

Public Function FindNearestItem(ByVal X As Single, ByVal Y As Single, Optional ByVal Direction As LvwFindDirectionConstants) As LvwListItem
Attribute FindNearestItem.VB_Description = "Finds an item nearest to the position specified and returns a reference to that item."
Select Case Direction
    Case LvwFindDirectionUndefined, LvwFindDirectionPrior, LvwFindDirectionNext, LvwFindDirectionEnd, LvwFindDirectionHome, LvwFindDirectionLeft, LvwFindDirectionUp, LvwFindDirectionRight, LvwFindDirectionDown
    Case Else
        Err.Raise 380
End Select
If ListViewHandle <> 0 Then
    Dim LVFI As LVFINDINFO, Index As Long
    With LVFI
    .PT.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    .PT.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    .VKDirection = Direction
    .Flags = LVFI_NEARESTXY
    End With
    Index = SendMessage(ListViewHandle, LVM_FINDITEM, -1, ByVal VarPtr(LVFI))
    If Index > -1 Then
        If PropVirtualMode = False Then
            Set FindNearestItem = Me.ListItems(Index + 1)
        Else
            Set FindNearestItem = New LvwListItem
            FindNearestItem.FInit ObjPtr(Me), Index + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
        End If
    End If
End If
End Function

Public Function FindSubItem(ByVal Text As String, Optional ByVal Index As Long, Optional ByRef SubItemIndex As Long, Optional ByVal Partial As Boolean, Optional ByVal Wrap As Boolean) As LvwListItem
Attribute FindSubItem.VB_Description = "Finds a sub item in the list and returns a reference to that item."
If PropVirtualMode = True Then Err.Raise Number:=5, Description:="This functionality is disabled when virtual mode is on."
If ListViewHandle <> 0 Then
    If Index >= 0 And SubItemIndex >= 0 Then
        Dim Count As Long, SubItemCount As Long
        Count = SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&)
        SubItemCount = Me.ColumnHeaders.Count - 1 ' Deduct 1 for SubItem 0
        If Count > 0 And SubItemCount > 0 Then
            If Index <= Count And SubItemIndex <= SubItemCount Then
                If Index = 0 Then Index = 1
                Dim i As Long
                If SubItemIndex > 0 Then
                    If Partial = False Then
                        For i = Index To Count
                            If StrComp(Me.FListItemText(i, SubItemIndex), Text, vbTextCompare) = 0 Then
                                Set FindSubItem = Me.ListItems(i)
                                Exit For
                            End If
                        Next i
                    Else
                        For i = Index To Count
                            If InStr(1, Me.FListItemText(i, SubItemIndex), Text, vbTextCompare) > 0 Then
                                Set FindSubItem = Me.ListItems(i)
                                Exit For
                            End If
                        Next i
                    End If
                    If FindSubItem Is Nothing And Wrap = True Then
                        If Partial = False Then
                            For i = 1 To (Index - 1)
                                If StrComp(Me.FListItemText(i, SubItemIndex), Text, vbTextCompare) = 0 Then
                                    Set FindSubItem = Me.ListItems(i)
                                    Exit For
                                End If
                            Next i
                        Else
                            For i = 1 To (Index - 1)
                                If InStr(1, Me.FListItemText(i, SubItemIndex), Text, vbTextCompare) > 0 Then
                                    Set FindSubItem = Me.ListItems(i)
                                    Exit For
                                End If
                            Next i
                        End If
                    End If
                Else
                    Dim j As Long
                    If Partial = False Then
                        For i = Index To Count
                            For j = 1 To SubItemCount
                                If StrComp(Me.FListItemText(i, j), Text, vbTextCompare) = 0 Then
                                    Set FindSubItem = Me.ListItems(i)
                                    SubItemIndex = j
                                    Exit For
                                End If
                            Next j
                            If SubItemIndex > 0 Then Exit For
                        Next i
                    Else
                        For i = Index To Count
                            For j = 1 To SubItemCount
                                If InStr(1, Me.FListItemText(i, j), Text, vbTextCompare) > 0 Then
                                    Set FindSubItem = Me.ListItems(i)
                                    SubItemIndex = j
                                    Exit For
                                End If
                            Next j
                            If SubItemIndex > 0 Then Exit For
                        Next i
                    End If
                    If FindSubItem Is Nothing And Wrap = True Then
                        If Partial = False Then
                            For i = 1 To (Index - 1)
                                For j = 1 To SubItemCount
                                    If StrComp(Me.FListItemText(i, j), Text, vbTextCompare) = 0 Then
                                        Set FindSubItem = Me.ListItems(i)
                                        SubItemIndex = j
                                        Exit For
                                    End If
                                Next j
                                If SubItemIndex > 0 Then Exit For
                            Next i
                        Else
                            For i = 1 To (Index - 1)
                                For j = 1 To SubItemCount
                                    If InStr(1, Me.FListItemText(i, j), Text, vbTextCompare) > 0 Then
                                        Set FindSubItem = Me.ListItems(i)
                                        SubItemIndex = j
                                        Exit For
                                    End If
                                Next j
                                If SubItemIndex > 0 Then Exit For
                            Next i
                        End If
                    End If
                End If
            Else
                Err.Raise 380
            End If
        End If
    Else
        Err.Raise 380
    End If
End If
End Function

Public Function GetVisibleCount() As Long
Attribute GetVisibleCount.VB_Description = "Returns the number of fully visible list items. If the list view is in 'icon', 'small icon' or 'tile' view then the return value is the total number of list items."
If ListViewHandle <> 0 Then GetVisibleCount = SendMessage(ListViewHandle, LVM_GETCOUNTPERPAGE, 0, ByVal 0&)
End Function

Public Function GetSelectedCount() As Long
Attribute GetSelectedCount.VB_Description = "Returns the number of selected items."
If ListViewHandle <> 0 Then GetSelectedCount = SendMessage(ListViewHandle, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
End Function

Public Function GetHeaderHeight() As Single
Attribute GetHeaderHeight.VB_Description = "Retrieves the height of the header control in 'report' view."
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        Dim RC As RECT
        GetWindowRect ListViewHeaderHandle, RC
        GetHeaderHeight = UserControl.ScaleY((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
    End If
End If
End Function

Public Property Get TopItem() As LvwListItem
Attribute TopItem.VB_Description = "Returns a reference to the topmost visible list item."
Attribute TopItem.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    If SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&) > 0 Then
        Dim RC As RECT, iItem As Long
        Select Case PropView
            Case LvwViewReport
                If PropGroupView = False Or ComCtlsSupportLevel() = 0 Then
                    If PropVirtualMode = False Then
                        Set TopItem = PtrToObj(Me.FListItemPtr(SendMessage(ListViewHandle, LVM_GETTOPINDEX, 0, ByVal 0&) + 1))
                    Else
                        Set TopItem = New LvwListItem
                        TopItem.FInit ObjPtr(Me), SendMessage(ListViewHandle, LVM_GETTOPINDEX, 0, ByVal 0&) + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                    End If
                ElseIf ComCtlsSupportLevel() >= 2 Then
                    ' Not supported if ComCtlsSupportLevel() = 1 and group view property is set to true.
                    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
                    If ListViewHeaderHandle <> 0 Then
                        Dim WndRect As RECT, LVI_V60 As LVITEM_V60, SubsetCount As Long, LastGroupID As Long
                        GetWindowRect ListViewHeaderHandle, WndRect
                        iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_ALL Or LVNI_VISIBLEORDER)
                        Do While iItem > -1
                            If PropGroupSubsetCount > 0 Then
                                With LVI_V60
                                .LVI.Mask = LVIF_GROUPID
                                .LVI.iItem = iItem
                                SendMessage ListViewHandle, LVM_GETITEM, 0, ByVal VarPtr(LVI_V60)
                                If .iGroupId <> I_GROUPIDNONE Then
                                    If .iGroupId <> LastGroupID Then SubsetCount = 0
                                    If Me.FGroupSubseted(.iGroupId) = True Then
                                        SubsetCount = SubsetCount + 1
                                    Else
                                        SubsetCount = 0
                                    End If
                                Else
                                    SubsetCount = 0
                                End If
                                LastGroupID = .iGroupId
                                End With
                            Else
                                SubsetCount = 0
                            End If
                            If SubsetCount <= PropGroupSubsetCount Then
                                SetRect RC, LVIR_BOUNDS, 0, 0, 0
                                SendMessage ListViewHandle, LVM_GETITEMRECT, iItem, ByVal VarPtr(RC)
                                If RC.Top >= (WndRect.Bottom - WndRect.Top) Then
                                    Set TopItem = PtrToObj(Me.FListItemPtr(iItem + 1))
                                    Exit Do
                                End If
                            End If
                            iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItem, ByVal LVNI_ALL Or LVNI_VISIBLEORDER)
                        Loop
                    End If
                End If
            Case LvwViewList
                ' LVM_GETTOPINDEX works here in all scenarios.
                If PropVirtualMode = False Then
                    Set TopItem = PtrToObj(Me.FListItemPtr(SendMessage(ListViewHandle, LVM_GETTOPINDEX, 0, ByVal 0&) + 1))
                Else
                    Set TopItem = New LvwListItem
                    TopItem.FInit ObjPtr(Me), SendMessage(ListViewHandle, LVM_GETTOPINDEX, 0, ByVal 0&) + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                End If
            Case Else
                If PropGroupView = False Or ComCtlsSupportLevel() = 0 Then
                    ' Not supported if ComCtlsSupportLevel() >= 1 and group view property is set to true.
                    Dim LVRC As RECT, iItemResult As Long, Flags As Long
                    SendMessage ListViewHandle, LVM_GETVIEWRECT, 0, ByVal VarPtr(LVRC)
                    SetRect LVRC, 0, 0, (LVRC.Right - LVRC.Left), (LVRC.Bottom - LVRC.Top)
                    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_ALL)
                    iItemResult = -1
                    If ComCtlsSupportLevel() >= 2 Then Flags = LVNI_ALL Or LVNI_VISIBLEONLY Else Flags = LVNI_ALL
                    Do While iItem > -1
                        SetRect RC, LVIR_BOUNDS, 0, 0, 0
                        SendMessage ListViewHandle, LVM_GETITEMRECT, iItem, ByVal VarPtr(RC)
                        If RC.Right > LVRC.Left Then
                            If RC.Left < LVRC.Right Then
                                If RC.Bottom > LVRC.Top Then
                                    If RC.Top < LVRC.Bottom Then
                                        iItemResult = iItem
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                        iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItem, ByVal Flags)
                    Loop
                    If iItemResult > -1 Then
                        ' Now try to move top-left to get the topmost visible list item.
                        Dim iItemTemp As Long
                        iItemTemp = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItemResult, ByVal LVNI_ALL Or LVNI_ABOVE)
                        If iItemTemp > -1 Then
                            SetRect RC, LVIR_BOUNDS, 0, 0, 0
                            SendMessage ListViewHandle, LVM_GETITEMRECT, iItemTemp, ByVal VarPtr(RC)
                            If RC.Right <= LVRC.Left Or RC.Left >= LVRC.Right Or RC.Bottom <= LVRC.Top Or RC.Top >= LVRC.Bottom Then iItemTemp = -1
                        End If
                        If iItemTemp = -1 Then
                            iItemTemp = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItemResult, ByVal LVNI_ALL Or LVNI_TOLEFT)
                            SetRect RC, LVIR_BOUNDS, 0, 0, 0
                            SendMessage ListViewHandle, LVM_GETITEMRECT, iItemTemp, ByVal VarPtr(RC)
                            If RC.Right <= LVRC.Left Or RC.Left >= LVRC.Right Or RC.Bottom <= LVRC.Top Or RC.Top >= LVRC.Bottom Then iItemTemp = -1
                        End If
                        iItem = iItemTemp
                        Do While iItem > -1
                            SetRect RC, LVIR_BOUNDS, 0, 0, 0
                            SendMessage ListViewHandle, LVM_GETITEMRECT, iItem, ByVal VarPtr(RC)
                            If RC.Right > LVRC.Left Then
                                If RC.Left < LVRC.Right Then
                                    If RC.Bottom > LVRC.Top Then
                                        If RC.Top < LVRC.Bottom Then
                                            iItemResult = iItem
                                        End If
                                    End If
                                End If
                            End If
                            iItemTemp = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItem, ByVal LVNI_ALL Or LVNI_ABOVE)
                            If iItemTemp > -1 Then
                                SetRect RC, LVIR_BOUNDS, 0, 0, 0
                                SendMessage ListViewHandle, LVM_GETITEMRECT, iItemTemp, ByVal VarPtr(RC)
                                If RC.Right <= LVRC.Left Or RC.Left >= LVRC.Right Or RC.Bottom <= LVRC.Top Or RC.Top >= LVRC.Bottom Then iItemTemp = -1
                            End If
                            If iItemTemp = -1 Then
                                iItemTemp = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItem, ByVal LVNI_ALL Or LVNI_TOLEFT)
                                If iItemTemp > -1 Then
                                    SetRect RC, LVIR_BOUNDS, 0, 0, 0
                                    SendMessage ListViewHandle, LVM_GETITEMRECT, iItemTemp, ByVal VarPtr(RC)
                                    If RC.Right <= LVRC.Left Or RC.Left >= LVRC.Right Or RC.Bottom <= LVRC.Top Or RC.Top >= LVRC.Bottom Then iItemTemp = -1
                                End If
                            End If
                            iItem = iItemTemp
                        Loop
                        Set TopItem = PtrToObj(Me.FListItemPtr(iItemResult + 1))
                        If PropVirtualMode = False Then
                            Set TopItem = PtrToObj(Me.FListItemPtr(iItemResult + 1))
                        Else
                            Set TopItem = New LvwListItem
                            TopItem.FInit ObjPtr(Me), iItemResult + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                        End If
                    End If
                End If
        End Select
    End If
End If
End Property

Public Property Get SelectedItem() As LvwListItem
Attribute SelectedItem.VB_Description = "Returns/sets a reference to the currently selected list item."
Attribute SelectedItem.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_ALL Or LVNI_FOCUSED)
    If iItem > -1 Then
        If PropVirtualMode = False Then
            Set SelectedItem = Me.ListItems(iItem + 1)
        Else
            Set SelectedItem = New LvwListItem
            SelectedItem.FInit ObjPtr(Me), iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
        End If
    End If
End If
End Property

Public Property Let SelectedItem(ByVal Value As LvwListItem)
Set Me.SelectedItem = Value
End Property

Public Property Set SelectedItem(ByVal Value As LvwListItem)
If ListViewHandle <> 0 Then
    If Not Value Is Nothing Then
        Value.Selected = True
    Else
        Dim LVI As LVITEM
        With LVI
        .Mask = LVIF_STATE
        .StateMask = LVIS_FOCUSED
        .State = 0
        End With
        SendMessage ListViewHandle, LVM_SETITEMSTATE, -1, ByVal VarPtr(LVI)
        Call CheckItemFocus(0)
    End If
End If
End Property

Public Function SelectedIndices() As Collection
Attribute SelectedIndices.VB_Description = "Returns a reference to a collection containing the indexes to the selected items."
Set SelectedIndices = New Collection
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_ALL Or LVNI_SELECTED)
    Do While iItem > -1
        SelectedIndices.Add (iItem + 1)
        iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItem, ByVal LVNI_ALL Or LVNI_SELECTED)
    Loop
End If
End Function

Public Function GhostedIndices() As Collection
Attribute GhostedIndices.VB_Description = "Returns a reference to a collection containing the indexes to the ghosted items."
Err.Raise Number:=91, Description:="This functionality is disabled when virtual mode is on."
Set GhostedIndices = New Collection
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_ALL Or LVNI_CUT)
    Do While iItem > -1
        GhostedIndices.Add (iItem + 1)
        iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, iItem, ByVal LVNI_ALL Or LVNI_CUT)
    Loop
End If
End Function

Public Function CheckedIndices() As Collection
Attribute CheckedIndices.VB_Description = "Returns a reference to a collection containing the indexes to the checked items."
Set CheckedIndices = New Collection
If ListViewHandle <> 0 Then
    Dim iItem As Long
    For iItem = 0 To (SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&) - 1)
        If StateImageMaskToIndex(SendMessage(ListViewHandle, LVM_GETITEMSTATE, iItem, ByVal LVIS_STATEIMAGEMASK) And LVIS_STATEIMAGEMASK) = IIL_CHECKED Then CheckedIndices.Add (iItem + 1)
    Next iItem
End If
End Function

Public Property Get HotItem() As LvwListItem
Attribute HotItem.VB_Description = "Returns/sets a reference to the currently hot list item. This is only meaningful if the hot tracking property is set to true."
Attribute HotItem.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETHOTITEM, 0, ByVal 0&)
    If iItem > -1 Then
        If PropVirtualMode = False Then
            Set HotItem = Me.ListItems(iItem + 1)
        Else
            Set HotItem = New LvwListItem
            HotItem.FInit ObjPtr(Me), iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
        End If
    End If
End If
End Property

Public Property Let HotItem(ByVal Value As LvwListItem)
Set Me.HotItem = Value
End Property

Public Property Set HotItem(ByVal Value As LvwListItem)
If ListViewHandle <> 0 Then
    If Not Value Is Nothing Then
        Value.Hot = True
    Else
        SendMessage ListViewHandle, LVM_SETHOTITEM, -1, ByVal 0&
    End If
End If
End Property

Public Property Get SelectedColumn() As LvwColumnHeader
Attribute SelectedColumn.VB_Description = "Returns/sets a reference to the currently selected column. Requires comctl32.dll version 6.0 or higher."
Attribute SelectedColumn.VB_MemberFlags = "400"
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    Dim Index As Long
    Index = SendMessage(ListViewHandle, LVM_GETSELECTEDCOLUMN, 0, ByVal 0&)
    If Index > -1 Then Set SelectedColumn = Me.ColumnHeaders(Index + 1)
End If
End Property

Public Property Let SelectedColumn(ByVal Value As LvwColumnHeader)
Set Me.SelectedColumn = Value
End Property

Public Property Set SelectedColumn(ByVal Value As LvwColumnHeader)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 1 Then
    If Value Is Nothing Then
        SendMessage ListViewHandle, LVM_SETSELECTEDCOLUMN, -1, ByVal 0&
    Else
        If Not PropPicture Is Nothing Then If PropPictureAlignment = LvwPictureAlignmentTile And PropPictureWatermark = False Then Exit Property
        SendMessage ListViewHandle, LVM_SETSELECTEDCOLUMN, Value.Index - 1, ByVal 0&
    End If
End If
End Property

Public Property Get SelectionMark() As LvwListItem
Attribute SelectionMark.VB_Description = "Returns/sets the selection mark. A selection mark is that list item from which a multiple selection starts."
Attribute SelectionMark.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETSELECTIONMARK, 0, ByVal 0&)
    If iItem > -1 Then
        If PropVirtualMode = False Then
            Set SelectionMark = Me.ListItems(iItem + 1)
        Else
            Set SelectionMark = New LvwListItem
            SelectionMark.FInit ObjPtr(Me), iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
        End If
    End If
End If
End Property

Public Property Let SelectionMark(ByVal Value As LvwListItem)
Set Me.SelectionMark = Value
End Property

Public Property Set SelectionMark(ByVal Value As LvwListItem)
If ListViewHandle <> 0 Then
    If Not Value Is Nothing Then
        Dim iItem As Long
        iItem = Value.Index - 1
        SendMessage ListViewHandle, LVM_SETSELECTIONMARK, 0, ByVal iItem
    Else
        SendMessage ListViewHandle, LVM_SETSELECTIONMARK, 0, ByVal -1&
    End If
End If
End Property

Public Property Get ColumnWidth() As Single
Attribute ColumnWidth.VB_Description = "Returns/sets the width of a column in 'list' view."
Attribute ColumnWidth.VB_MemberFlags = "400"
If PropView = LvwViewList Then
    If ListViewHandle <> 0 Then ColumnWidth = UserControl.ScaleX(SendMessage(ListViewHandle, LVM_GETCOLUMNWIDTH, 0, ByVal 0&), vbPixels, vbContainerSize)
Else
    Err.Raise Number:=394, Description:="Get supported in 'list' view only"
End If
End Property

Public Property Let ColumnWidth(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim LngValue As Long
LngValue = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If LngValue > 0 Then
    If PropView = LvwViewList Then
        ListViewMemoryColumnWidth = LngValue
        If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_SETCOLUMNWIDTH, 0, ByVal LngValue
    Else
        Err.Raise Number:=383, Description:="Set supported in 'list' view only"
    End If
Else
    Err.Raise 380
End If
End Property

Public Property Get IconSpacingWidth() As Single
Attribute IconSpacingWidth.VB_Description = "Returns/sets the spacing width between icons in 'icon' and 'small icon' view."
Attribute IconSpacingWidth.VB_MemberFlags = "400"
If PropView = LvwViewIcon Then
    If ListViewHandle <> 0 Then IconSpacingWidth = UserControl.ScaleX(LoWord(SendMessage(ListViewHandle, LVM_GETITEMSPACING, 0, ByVal 0&)), vbPixels, vbContainerSize)
ElseIf PropView = LvwViewSmallIcon Then
    If ListViewHandle <> 0 Then IconSpacingWidth = UserControl.ScaleX(LoWord(SendMessage(ListViewHandle, LVM_GETITEMSPACING, 1, ByVal 0&)), vbPixels, vbContainerSize)
Else
    Err.Raise Number:=394, Description:="Get supported in 'icon' and 'small icon' view only"
End If
End Property

Public Property Let IconSpacingWidth(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim LngValueX As Long, LngValueY As Long
LngValueX = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If LngValueX >= 0 Then
    If PropView = LvwViewIcon Then
        If ListViewHandle <> 0 Then
            LngValueY = HiWord(SendMessage(ListViewHandle, LVM_GETITEMSPACING, 0, ByVal 0&))
            SendMessage ListViewHandle, LVM_SETICONSPACING, 0, ByVal MakeDWord(LngValueX, LngValueY)
            Me.Refresh
        End If
    Else
        Err.Raise Number:=383, Description:="Set supported in 'icon' view only"
    End If
Else
    Err.Raise 380
End If
End Property

Public Property Get IconSpacingHeight() As Single
Attribute IconSpacingHeight.VB_Description = "Returns/sets the spacing height between icons in 'icon' and 'small icon' view."
Attribute IconSpacingHeight.VB_MemberFlags = "400"
If PropView = LvwViewIcon Then
    If ListViewHandle <> 0 Then IconSpacingHeight = UserControl.ScaleY(HiWord(SendMessage(ListViewHandle, LVM_GETITEMSPACING, 0, ByVal 0&)), vbPixels, vbContainerSize)
ElseIf PropView = LvwViewSmallIcon Then
    If ListViewHandle <> 0 Then IconSpacingHeight = UserControl.ScaleY(HiWord(SendMessage(ListViewHandle, LVM_GETITEMSPACING, 1, ByVal 0&)), vbPixels, vbContainerSize)
Else
    Err.Raise Number:=394, Description:="Get supported in 'icon' and 'small icon' view only"
End If
End Property

Public Property Let IconSpacingHeight(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim LngValueX As Long, LngValueY As Long
LngValueY = CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
If LngValueY >= 0 Then
    If PropView = LvwViewIcon Then
        If ListViewHandle <> 0 Then
            LngValueX = LoWord(SendMessage(ListViewHandle, LVM_GETITEMSPACING, 0, ByVal 0&))
            SendMessage ListViewHandle, LVM_SETICONSPACING, 0, ByVal MakeDWord(LngValueX, LngValueY)
            Me.Refresh
        End If
    Else
        Err.Raise Number:=383, Description:="Set supported in 'icon' view only"
    End If
Else
    Err.Raise 380
End If
End Property

Public Sub ResetIconSpacing()
Attribute ResetIconSpacing.VB_Description = "Resets the spacing between icons to the default spacing width and height in 'icon' view."
If ListViewHandle <> 0 Then
    SendMessage ListViewHandle, LVM_SETICONSPACING, 0, ByVal -1&
    Me.Refresh
End If
End Sub

Public Sub StartLabelEdit()
Attribute StartLabelEdit.VB_Description = "Begins a label editing operation on a list item. This method will fail if the label edit property is set to disabled."
If ListViewHandle <> 0 Then
    ListViewStartLabelEdit = True
    SendMessage ListViewHandle, LVM_EDITLABEL, ListViewFocusIndex - 1, ByVal 0&
    ListViewStartLabelEdit = False
End If
End Sub

Public Sub EndLabelEdit()
Attribute EndLabelEdit.VB_Description = "Ends the label editing operation on a list item."
If ListViewHandle <> 0 Then
    If ComCtlsSupportLevel() >= 1 Then
        SendMessage ListViewHandle, LVM_CANCELEDITLABEL, 0, ByVal 0&
    Else
        SendMessage ListViewHandle, LVM_EDITLABEL, -1, ByVal 0&
    End If
End If
End Sub

Public Sub Scroll(ByVal X As Single, ByVal Y As Single)
Attribute Scroll.VB_Description = "Scrolls the content. When the list view is in 'report' view, the X and Y arguments will be rounded up to the nearest number that form a whole line increment."
If ListViewHandle <> 0 Then SendMessage ListViewHandle, LVM_SCROLL, CLng(UserControl.ScaleX(X, vbContainerSize, vbPixels)), ByVal CLng(UserControl.ScaleX(Y, vbContainerSize, vbPixels))
End Sub

Public Sub ResetEmptyMarkup()
Attribute ResetEmptyMarkup.VB_Description = "Method to force the control to request again for a markup text. Requires comctl32.dll version 6.1 or higher."
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then SendMessage ListViewHandle, LVM_RESETEMPTYTEXT, 0, ByVal 0&
End Sub

Public Sub ComputeControlSize(ByVal VisibleCount As Long, ByRef Width As Single, ByRef Height As Single, Optional ByVal ProposedWidth As Single, Optional ByVal ProposedHeight As Single)
Attribute ComputeControlSize.VB_Description = "A method that returns the width and height for a given number of visible list items."
If VisibleCount < 0 Then Err.Raise 380
If ListViewHandle <> 0 Then
    Dim RetVal As Long, RC(0 To 1) As RECT, ProposedX As Long, ProposedY As Long
    GetWindowRect ListViewHandle, RC(0)
    GetClientRect ListViewHandle, RC(1)
    With UserControl
    If ProposedWidth <> 0 Then
        ProposedX = CLng(.ScaleX(ProposedWidth, vbContainerSize, vbPixels))
    Else
        ProposedX = -1
    End If
    If ProposedHeight <> 0 Then
        ProposedY = CLng(.ScaleY(ProposedHeight, vbContainerSize, vbPixels))
    Else
        ProposedY = -1
    End If
    RetVal = SendMessage(ListViewHandle, LVM_APPROXIMATEVIEWRECT, IIf(PropView = LvwViewReport, VisibleCount - 1, VisibleCount), MakeDWord(ProposedX, ProposedY))
    If LoWord(RetVal) <> 0 Then Width = .ScaleX(LoWord(RetVal) + ((RC(0).Right - RC(0).Left) - (RC(1).Right - RC(1).Left)), vbPixels, vbContainerSize)
    If HiWord(RetVal) <> 0 Then Height = .ScaleY(HiWord(RetVal) + ((RC(0).Bottom - RC(0).Top) - (RC(1).Bottom - RC(1).Top)), vbPixels, vbContainerSize)
    End With
End If
End Sub

Public Function TextWidth(ByVal Text As String) As Single
Attribute TextWidth.VB_Description = "Returns the text width of the given string using the current font of the list view."
If ListViewHandle <> 0 Then
    Dim Pixels As Long
    Pixels = SendMessage(ListViewHandle, LVM_GETSTRINGWIDTH, 0, ByVal StrPtr(Text))
    If Pixels > 0 Then TextWidth = UserControl.ScaleX(Pixels, vbPixels, vbContainerSize)
End If
End Function

Public Property Get DropHighlight() As LvwListItem
Attribute DropHighlight.VB_Description = "Returns/sets a reference to a list item and highlights it with the system highlight color."
Attribute DropHighlight.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim iItem As Long
    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_ALL Or LVNI_DROPHILITED)
    If iItem > -1 Then
        If PropVirtualMode = False Then
            Set DropHighlight = Me.ListItems(iItem + 1)
        Else
            Set DropHighlight = New LvwListItem
            DropHighlight.FInit ObjPtr(Me), iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
        End If
    End If
End If
End Property

Public Property Let DropHighlight(ByVal Value As LvwListItem)
Set Me.DropHighlight = Value
End Property

Public Property Set DropHighlight(ByVal Value As LvwListItem)
If ListViewHandle <> 0 Then
    Dim iItem As Long, LVI As LVITEM
    LVI.StateMask = LVIS_DROPHILITED
    If Not Value Is Nothing Then
        iItem = Value.Index - 1
        If iItem <> SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_ALL Or LVNI_DROPHILITED) Then
            With LVI
            .State = 0
            SendMessage ListViewHandle, LVM_SETITEMSTATE, -1, ByVal VarPtr(LVI)
            If iItem > -1 Then
                .State = LVIS_DROPHILITED
                SendMessage ListViewHandle, LVM_SETITEMSTATE, iItem, ByVal VarPtr(LVI)
            End If
            End With
        End If
    Else
        LVI.State = 0
        SendMessage ListViewHandle, LVM_SETITEMSTATE, -1, ByVal VarPtr(LVI)
    End If
End If
End Property

Public Property Get InsertMark(Optional ByRef After As Boolean) As LvwListItem
Attribute InsertMark.VB_Description = "Returns/sets a reference to a list item where an insertion mark is positioned. Requires comctl32.dll version 6.1 or higher."
Attribute InsertMark.VB_MemberFlags = "400"
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVIM As LVINSERTMARK
    With LVIM
    .cbSize = LenB(LVIM)
    SendMessage ListViewHandle, LVM_GETINSERTMARK, 0, ByVal VarPtr(LVIM)
    If .iItem > -1 Then
        If PropVirtualMode = False Then
            Set InsertMark = Me.ListItems(.iItem + 1)
        Else
            Set InsertMark = New LvwListItem
            InsertMark.FInit ObjPtr(Me), .iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
        End If
        After = CBool((CDbl(.dwFlags) - CDbl(vbDropEffectScroll)) = LVIM_AFTER)
    End If
    End With
End If
End Property

Public Property Let InsertMark(Optional ByRef After As Boolean, ByVal Value As LvwListItem)
Set Me.InsertMark(After) = Value
End Property

Public Property Set InsertMark(Optional ByRef After As Boolean, ByVal Value As LvwListItem)
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim LVIM As LVINSERTMARK
    With LVIM
    .cbSize = LenB(LVIM)
    If Value Is Nothing Then
        .iItem = -1
        .dwFlags = 0
    Else
        .iItem = Value.Index - 1
        .dwFlags = IIf(After = True, LVIM_AFTER, 0)
    End If
    End With
    SendMessage ListViewHandle, LVM_SETINSERTMARK, 0, ByVal VarPtr(LVIM)
End If
End Property

Public Property Get OLEDraggedItem() As LvwListItem
Attribute OLEDraggedItem.VB_Description = "Returns a reference to the currently dragged list item during an OLE drag/drop operation."
Attribute OLEDraggedItem.VB_MemberFlags = "400"
If ListViewDragIndex > 0 Then
    If PropVirtualMode = False Then
        Dim Ptr As Long
        Ptr = Me.FListItemPtr(ListViewDragIndex)
        If Ptr <> 0 Then Set OLEDraggedItem = PtrToObj(Ptr)
    Else
        Set OLEDraggedItem = New LvwListItem
        OLEDraggedItem.FInit ObjPtr(Me), ListViewDragIndex, vbNullString, 0, vbNullString, 0, 0, 0, 0
    End If
End If
End Property

Public Property Get SelectedGroup() As LvwGroup
Attribute SelectedGroup.VB_Description = "Returns/sets a reference to the currently selected group. Requires comctl32.dll version 6.1 or higher."
If PropVirtualMode = True Then Err.Raise Number:=5, Description:="This functionality is disabled when virtual mode is on."
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    Dim Index As Long
    Index = SendMessage(ListViewHandle, LVM_GETFOCUSEDGROUP, 0, ByVal 0&)
    If Index > -1 Then
        Dim LVG As LVGROUP
        With LVG
        .cbSize = LenB(LVG)
        .Mask = LVGF_GROUPID
        SendMessage ListViewHandle, LVM_GETGROUPINFOBYINDEX, Index, ByVal VarPtr(LVG)
        Dim Group As LvwGroup
        For Each Group In Me.Groups
            If Group.ID = .iGroupId Then
                Set SelectedGroup = Group
                Exit For
            End If
        Next Group
        End With
    End If
End If
End Property

Public Property Let SelectedGroup(ByVal Value As LvwGroup)
Set Me.SelectedGroup = Value
End Property

Public Property Set SelectedGroup(ByVal Value As LvwGroup)
If PropVirtualMode = True Then Err.Raise Number:=5, Description:="This functionality is disabled when virtual mode is on."
If ListViewHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
    If Not Value Is Nothing Then
        Value.Selected = True
    Else
        Dim Index As Long
        Index = SendMessage(ListViewHandle, LVM_GETFOCUSEDGROUP, 0, ByVal 0&)
        If Index > -1 Then
            Dim LVG As LVGROUP
            With LVG
            .cbSize = LenB(LVG)
            SendMessage ListViewHandle, LVM_GETGROUPINFOBYINDEX, Index, ByVal VarPtr(LVG)
            .Mask = LVGF_STATE
            .StateMask = LVGS_FOCUSED
            .State = 0
            SendMessage ListViewHandle, LVM_SETGROUPINFO, .iGroupId, ByVal VarPtr(LVG)
            End With
        End If
    End If
End If
End Property

Public Property Get ColumnOrder() As Variant
Attribute ColumnOrder.VB_Description = "Returns/sets the column order of the list view. All the position indexes are zero-based."
Attribute ColumnOrder.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    Dim Count As Long
    Count = Me.ColumnHeaders.Count
    If Count > 0 Then
        Dim ArgList() As Long
        ReDim ArgList(0 To (Count - 1)) As Long
        SendMessage ListViewHandle, LVM_GETCOLUMNORDERARRAY, Count, ByVal VarPtr(ArgList(0))
        ColumnOrder = ArgList()
    Else
        ColumnOrder = Empty
    End If
End If
End Property

Public Property Let ColumnOrder(ByVal ArgList As Variant)
If ListViewHandle <> 0 Then
    If IsArray(ArgList) Then
        Dim Ptr As Long
        CopyMemory Ptr, ByVal UnsignedAdd(VarPtr(ArgList), 8), 4
        If Ptr <> 0 Then
            Dim DimensionCount As Integer
            CopyMemory DimensionCount, ByVal Ptr, 2
            If DimensionCount = 1 Then
                Dim Arr() As Long, Count As Long, i As Long
                For i = LBound(ArgList) To UBound(ArgList)
                    Select Case VarType(ArgList(i))
                        Case vbLong, vbInteger, vbByte
                            If ArgList(i) >= 0 Then
                                ReDim Preserve Arr(0 To Count) As Long
                                Arr(Count) = ArgList(i)
                                Count = Count + 1
                            End If
                        Case vbDouble, vbSingle
                            If CLng(ArgList(i)) >= 0 Then
                                ReDim Preserve Arr(0 To Count) As Long
                                Arr(Count) = CLng(ArgList(i))
                                Count = Count + 1
                            End If
                    End Select
                Next i
                If Count > 0 Then
                    If SendMessage(ListViewHandle, LVM_SETCOLUMNORDERARRAY, Count, ByVal VarPtr(Arr(0))) = 0 Then Err.Raise 5
                End If
            Else
                Err.Raise Number:=5, Description:="Array must be single dimensioned"
            End If
        Else
            Err.Raise Number:=91, Description:="Array is not allocated"
        End If
    Else
        If Not IsEmpty(ArgList) Then Err.Raise 380
    End If
End If
End Property

Public Property Get ColumnFilterChangedTimeout() As Long
Attribute ColumnFilterChangedTimeout.VB_Description = "Returns/sets the time in milliseconds before the 'ColumnFilterChanged' event is fired afer a filter was changed. A value of -1 indicates that the 'ColumnFilterChanged' is fired only when the filter edit is completed."
Attribute ColumnFilterChangedTimeout.VB_MemberFlags = "400"
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then
        ' When passing zero in HDM_SETFILTERCHANGETIMEOUT the timeout interval will not be changed and returns the current value.
        ColumnFilterChangedTimeout = SendMessage(ListViewHeaderHandle, HDM_SETFILTERCHANGETIMEOUT, 0, ByVal 0&)
    End If
End If
End Property

Public Property Let ColumnFilterChangedTimeout(ByVal Value As Long)
If Value = 0 Or Value < -1 Then Err.Raise 380
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then SendMessage ListViewHeaderHandle, HDM_SETFILTERCHANGETIMEOUT, 0, ByVal Value
End If
End Property

Public Sub ResetForeColors()
Attribute ResetForeColors.VB_Description = "Resets the foreground color of particular list and list sub items that have been modified."
If PropVirtualMode = True Then Err.Raise Number:=5, Description:="This functionality is disabled when virtual mode is on."
If ListViewHandle <> 0 Then
    Dim ListItem As LvwListItem, i As Long
    SendMessage ListViewHandle, WM_SETREDRAW, 0, ByVal 0&
    For Each ListItem In Me.ListItems
        With ListItem
        .ForeColor = -1
        For i = 1 To .FListSubItemsCount
            .FListSubItemProp(i, 7) = -1
        Next i
        End With
    Next ListItem
    If PropRedraw = True Then SendMessage ListViewHandle, WM_SETREDRAW, 1, ByVal 0&
End If
End Sub

Private Sub SetVisualStylesHeader()
If ListViewHandle <> 0 Then
    If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles ListViewHeaderHandle
        Else
            RemoveVisualStyles ListViewHeaderHandle
        End If
    End If
End If
End Sub

Private Sub SetVisualStylesToolTip()
If ListViewHandle <> 0 Then
    If ListViewToolTipHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles ListViewToolTipHandle
        Else
            RemoveVisualStyles ListViewToolTipHandle
        End If
    End If
End If
End Sub

Private Sub SetVisualStylesHeaderToolTip()
If ListViewHandle <> 0 Then
    If ListViewHeaderToolTipHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles ListViewHeaderToolTipHandle
        Else
            RemoveVisualStyles ListViewHeaderToolTipHandle
        End If
    End If
End If
End Sub

Private Sub SetColumnsSubItemIndex(Optional ByVal CountOffset As Long)
If ListViewHandle = 0 Then Exit Sub
If (Me.ColumnHeaders.Count + CountOffset) > 0 Then
    Dim LVC As LVCOLUMN
    LVC.Mask = LVCF_SUBITEM
    Dim i As Long
    For i = 1 To (Me.ColumnHeaders.Count + CountOffset)
        LVC.iSubItem = i - 1
        SendMessage ListViewHandle, LVM_SETCOLUMN, i - 1, ByVal VarPtr(LVC)
    Next i
End If
End Sub

Private Sub SetColumnRTLReading(ByVal ColumnHeaderIndex As Long, ByVal Value As Boolean)
If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
If ListViewHeaderHandle = 0 Then Exit Sub
Dim HDI As HDITEM
With HDI
.Mask = HDI_FORMAT
SendMessage ListViewHeaderHandle, HDM_GETITEM, ColumnHeaderIndex - 1, ByVal VarPtr(HDI)
If Not Value = CBool((.fmt And HDF_RTLREADING) = HDF_RTLREADING) Then
    If Value = True Then
        If Not (.fmt And HDF_RTLREADING) = HDF_RTLREADING Then .fmt = .fmt Or HDF_RTLREADING
    Else
        If (.fmt And HDF_RTLREADING) = HDF_RTLREADING Then .fmt = .fmt And Not HDF_RTLREADING
    End If
    SendMessage ListViewHeaderHandle, HDM_SETITEM, ColumnHeaderIndex - 1, ByVal VarPtr(HDI)
End If
End With
End Sub

Private Sub RebuildListItems()
If PropVirtualMode = True Then Exit Sub
Dim Count As Long
Count = Me.ColumnHeaders.Count
If Count > 0 Then
    Dim ListItem As LvwListItem, i As Long, j As Long
    For Each ListItem In Me.ListItems
        i = i + 1
        With ListItem
        .Text = .Text
        If .FListSubItemsCount > 0 Then
            For j = 1 To Count
                If j <= .FListSubItemsCount Then
                    Me.FListItemText(i, j) = .FListSubItemProp(j, 3)
                Else
                    Me.FListItemText(i, j) = vbNullString
                End If
            Next j
        End If
        End With
    Next ListItem
    Me.Refresh
End If
End Sub

Private Sub CheckHeaderControl()
If ListViewHeaderHandle = 0 Then
    ListViewHeaderHandle = Me.hWndHeader
    If ListViewHeaderHandle <> 0 Then Call ComCtlsSetSubclass(ListViewHeaderHandle, Me, 4)
End If
If ListViewHeaderHandle <> 0 Then
    If Not PropColumnHeaderIconsName = "(None)" Then
        If PropColumnHeaderIconsControl Is Nothing Then
            Me.ColumnHeaderIcons = PropColumnHeaderIconsName
        End If
    End If
    Call SetVisualStylesHeader
    Me.AllowColumnCheckboxes = PropAllowColumnCheckboxes
    Me.ShowColumnTips = PropShowColumnTips
    Me.ClickableColumnHeaders = PropClickableColumnHeaders
    Me.HighlightColumnHeaders = PropHighlightColumnHeaders
    Me.TrackSizeColumnHeaders = PropTrackSizeColumnHeaders
    Me.ResizableColumnHeaders = PropResizableColumnHeaders
    Me.UseColumnChevron = PropUseColumnChevron
    Me.UseColumnFilterBar = PropUseColumnFilterBar
    Me.Refresh
End If
End Sub

Private Sub UpdateHeaderToolTipRect(ByVal hWnd As Long)
If ListViewHandle <> 0 And ListViewHeaderToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = hWnd
    .uId = 0
    GetClientRect hWnd, .RC
    SendMessage ListViewHeaderToolTipHandle, TTM_NEWTOOLRECT, 0, ByVal VarPtr(TI)
    End With
End If
End Sub

Private Sub CheckHeaderToolTipItem(ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long)
Static LastFlags As Long
If ListViewHandle <> 0 And ListViewHeaderToolTipHandle <> 0 Then
    Dim HDHTI As HDHITTESTINFO
    With HDHTI
    .PT.X = X
    .PT.Y = Y
    If SendMessage(hWnd, HDM_HITTEST, 0, ByVal VarPtr(HDHTI)) > -1 Then
        If (.Flags And HHT_ONDIVIDER) = 0 And (.Flags And HHT_ONDIVOPEN) = 0 And (.Flags And HHT_ONFILTER) = 0 Then
            If ListViewHeaderToolTipItem <> .iItem Then
                ListViewHeaderToolTipItem = .iItem
                SendMessage ListViewHeaderToolTipHandle, TTM_POP, 0, ByVal 0&
                LastFlags = 0
            Else
                If (.Flags And HHT_ONFILTERBUTTON) <> 0 Then
                    If (LastFlags And HHT_ONFILTERBUTTON) = 0 Then SendMessage ListViewHeaderToolTipHandle, TTM_POP, 0, ByVal 0&
                ElseIf (.Flags And HHT_ONDROPDOWN) <> 0 Then
                    If (LastFlags And HHT_ONDROPDOWN) = 0 Then SendMessage ListViewHeaderToolTipHandle, TTM_POP, 0, ByVal 0&
                Else
                    If (LastFlags And HHT_ONFILTERBUTTON) <> 0 Or (LastFlags And HHT_ONDROPDOWN) <> 0 Then SendMessage ListViewHeaderToolTipHandle, TTM_POP, 0, ByVal 0&
                End If
                LastFlags = .Flags
            End If
        Else
            ListViewHeaderToolTipItem = -1
            If ListViewHeaderToolTipHandle <> 0 Then SendMessage ListViewHeaderToolTipHandle, TTM_POP, 0, ByVal 0&
            LastFlags = 0
        End If
    Else
        ListViewHeaderToolTipItem = -1
        If ListViewHeaderToolTipHandle <> 0 Then SendMessage ListViewHeaderToolTipHandle, TTM_POP, 0, ByVal 0&
        LastFlags = 0
    End If
    End With
Else
    LastFlags = 0
End If
End Sub

Private Function GetColumnToolTipText(ByVal hWnd As Long, ByVal Pos As Long) As String
Dim HDHTI As HDHITTESTINFO
With HDHTI
.PT.X = Get_X_lParam(Pos)
.PT.Y = Get_Y_lParam(Pos)
ScreenToClient hWnd, .PT
If SendMessage(hWnd, HDM_HITTEST, 0, ByVal VarPtr(HDHTI)) > -1 Then
    If (.Flags And HHT_ONDIVIDER) = 0 And (.Flags And HHT_ONDIVOPEN) = 0 And (.Flags And HHT_ONFILTER) = 0 Then
        If (.Flags And HHT_ONFILTERBUTTON) <> 0 Then
            GetColumnToolTipText = Me.ColumnHeaders(.iItem + 1).ToolTipTextFilterBtn
        ElseIf (.Flags And HHT_ONDROPDOWN) <> 0 Then
            GetColumnToolTipText = Me.ColumnHeaders(.iItem + 1).ToolTipTextDropDown
        Else
            GetColumnToolTipText = Me.ColumnHeaders(.iItem + 1).ToolTipText
        End If
    End If
End If
End With
End Function

Private Sub CheckItemFocus(ByVal Index As Long)
If ListViewHandle <> 0 Then
    Dim ParamValid As Boolean
    If PropVirtualMode = False Then
        ParamValid = CBool(Index > 0 And Index <= Me.ListItems.Count)
    Else
        ParamValid = CBool(Index > 0 And Index <= SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&))
    End If
    If Index <> ListViewFocusIndex Then
        ListViewFocusIndex = Index
        If ParamValid = True Then
            Dim ListItem As LvwListItem
            If PropVirtualMode = False Then
                Set ListItem = Me.ListItems(ListViewFocusIndex)
            Else
                Set ListItem = New LvwListItem
                ListItem.FInit ObjPtr(Me), ListViewFocusIndex, vbNullString, 0, vbNullString, 0, 0, 0, 0
            End If
            RaiseEvent ItemFocus(ListItem)
        End If
    End If
End If
End Sub

Private Sub SortListItems()
If PropVirtualMode = True Then Exit Sub
If ListViewHandle <> 0 Then
    If SendMessage(ListViewHandle, LVM_GETITEMCOUNT, 0, ByVal 0&) > 0 Then
        If PropSortKey > Me.ColumnHeaders.Count Then PropSortKey = Me.ColumnHeaders.Count
        Dim Address As Long
        Select Case PropSortType
            Case LvwSortTypeBinary
                Address = ProcPtr(AddressOf ComCtlsLvwSortingFunctionBinary)
            Case LvwSortTypeText
                Address = ProcPtr(AddressOf ComCtlsLvwSortingFunctionText)
            Case LvwSortTypeNumeric
                Address = ProcPtr(AddressOf ComCtlsLvwSortingFunctionNumeric)
            Case LvwSortTypeCurrency
                Address = ProcPtr(AddressOf ComCtlsLvwSortingFunctionCurrency)
            Case LvwSortTypeDate
                Address = ProcPtr(AddressOf ComCtlsLvwSortingFunctionDate)
            Case LvwSortTypeLogical
                Address = ProcPtr(AddressOf ComCtlsLvwSortingFunctionLogical)
        End Select
        If Address <> 0 Then
            Dim This As ISubclass
            Set This = Me
            SendMessage ListViewHandle, LVM_SORTITEMSEX, ObjPtr(This), ByVal Address
        End If
    End If
End If
End Sub

Private Function ListItemsSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
ListItemsSortingFunctionBinary = lstrcmp(StrPtr(Text1), StrPtr(Text2))
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionBinary = -ListItemsSortingFunctionBinary
End Function

Private Function ListItemsSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
ListItemsSortingFunctionText = lstrcmpi(StrPtr(Text1), StrPtr(Text2))
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionText = -ListItemsSortingFunctionText
End Function

Private Function ListItemsSortingFunctionNumeric(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
Dim Dbl1 As Double, Dbl2 As Double
On Error Resume Next
Dbl1 = CDbl(Text1)
Dbl2 = CDbl(Text2)
On Error GoTo 0
ListItemsSortingFunctionNumeric = Sgn(Dbl1 - Dbl2)
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionNumeric = -ListItemsSortingFunctionNumeric
End Function

Private Function ListItemsSortingFunctionCurrency(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
Dim Cur1 As Currency, Cur2 As Currency
On Error Resume Next
Cur1 = CCur(Text1)
Cur2 = CCur(Text2)
On Error GoTo 0
ListItemsSortingFunctionCurrency = Sgn(Cur1 - Cur2)
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionCurrency = -ListItemsSortingFunctionCurrency
End Function

Private Function ListItemsSortingFunctionDate(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
Dim Date1 As Date, Date2 As Date
On Error Resume Next
Date1 = CDate(Text1)
Date2 = CDate(Text2)
On Error GoTo 0
ListItemsSortingFunctionDate = Sgn(Date1 - Date2)
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionDate = -ListItemsSortingFunctionDate
End Function

Private Function ListItemsSortingFunctionLogical(ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FListItemText(lParam1 + 1, PropSortKey)
Text2 = Me.FListItemText(lParam2 + 1, PropSortKey)
ListItemsSortingFunctionLogical = StrCmpLogical(StrPtr(Text1), StrPtr(Text2))
If PropSortOrder = LvwSortOrderDescending Then ListItemsSortingFunctionLogical = -ListItemsSortingFunctionLogical
End Function

Private Function GroupsSortingFunctionBinary(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As LvwSortOrderConstants) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FGroupHeader(lParam1)
Text2 = Me.FGroupHeader(lParam2)
GroupsSortingFunctionBinary = lstrcmp(StrPtr(Text1), StrPtr(Text2))
If SortOrder = LvwSortOrderDescending Then GroupsSortingFunctionBinary = -GroupsSortingFunctionBinary
End Function

Private Function GroupsSortingFunctionText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As LvwSortOrderConstants) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FGroupHeader(lParam1)
Text2 = Me.FGroupHeader(lParam2)
GroupsSortingFunctionText = lstrcmpi(StrPtr(Text1), StrPtr(Text2))
If SortOrder = LvwSortOrderDescending Then GroupsSortingFunctionText = -GroupsSortingFunctionText
End Function

Private Function GroupsSortingFunctionNumeric(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As LvwSortOrderConstants) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FGroupHeader(lParam1)
Text2 = Me.FGroupHeader(lParam2)
Dim Dbl1 As Double, Dbl2 As Double
On Error Resume Next
Dbl1 = CDbl(Text1)
Dbl2 = CDbl(Text2)
On Error GoTo 0
GroupsSortingFunctionNumeric = Sgn(Dbl1 - Dbl2)
If SortOrder = LvwSortOrderDescending Then GroupsSortingFunctionNumeric = -GroupsSortingFunctionNumeric
End Function

Private Function GroupsSortingFunctionCurrency(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As LvwSortOrderConstants) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FGroupHeader(lParam1)
Text2 = Me.FGroupHeader(lParam2)
Dim Cur1 As Currency, Cur2 As Currency
On Error Resume Next
Cur1 = CCur(Text1)
Cur2 = CCur(Text2)
On Error GoTo 0
GroupsSortingFunctionCurrency = Sgn(Cur1 - Cur2)
If SortOrder = LvwSortOrderDescending Then GroupsSortingFunctionCurrency = -GroupsSortingFunctionCurrency
End Function

Private Function GroupsSortingFunctionDate(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As LvwSortOrderConstants) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FGroupHeader(lParam1)
Text2 = Me.FGroupHeader(lParam2)
Dim Date1 As Date, Date2 As Date
On Error Resume Next
Date1 = CDate(Text1)
Date2 = CDate(Text2)
On Error GoTo 0
GroupsSortingFunctionDate = Sgn(Date1 - Date2)
If SortOrder = LvwSortOrderDescending Then GroupsSortingFunctionDate = -GroupsSortingFunctionDate
End Function

Private Function GroupsSortingFunctionLogical(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As LvwSortOrderConstants) As Long
Dim Text1 As String, Text2 As String
Text1 = Me.FGroupHeader(lParam1)
Text2 = Me.FGroupHeader(lParam2)
GroupsSortingFunctionLogical = StrCmpLogical(StrPtr(Text1), StrPtr(Text2))
If SortOrder = LvwSortOrderDescending Then GroupsSortingFunctionLogical = -GroupsSortingFunctionLogical
End Function

Private Function NextGroupID() As Long
Static ID As Long
ID = ID + 1
NextGroupID = ID
End Function

Private Function IsGroupAvailable(ByVal ID As Long) As Boolean
If ListViewHandle <> 0 Then IsGroupAvailable = CBool(SendMessage(ListViewHandle, LVM_HASGROUP, ID, ByVal 0&) <> 0)
End Function

Private Function GetGroupFromID(ByVal ID As Long) As LvwGroup
If IsGroupAvailable(ID) = True Then
    Dim Group As LvwGroup
    For Each Group In Me.Groups
        If Group.ID = ID Then
            Set GetGroupFromID = Group
            Exit For
        End If
    Next Group
End If
End Function

Private Sub GetWorkAreaRect(ByVal Index As Long, ByRef RC As RECT)
If ListViewHandle <> 0 Then
    Dim Count As Long
    SendMessage ListViewHandle, LVM_GETNUMBEROFWORKAREAS, 0, ByVal VarPtr(Count)
    If Count > 0 And Index <= Count And Index > 0 Then
        Dim ArrRC() As RECT
        ReDim ArrRC(1 To Count) As RECT
        SendMessage ListViewHandle, LVM_GETWORKAREAS, Count, ByVal VarPtr(ArrRC(1))
        LSet RC = ArrRC(Index)
    End If
End If
End Sub

Private Function GetFilterEditIndex(ByVal hWndFilterEdit As Long) As Long
If ListViewHandle = 0 Or hWndFilterEdit = 0 Then Exit Function
' If comctl32.dll version is 6.1 or higher then HDN_BEGINFILTEREDIT and HDN_ENDFILTEREDIT will be sent.
' Thus we return zero in order to not raise the events 'BeforeFilterEdit' and 'AfterFilterEdit' twice.
If ComCtlsSupportLevel() >= 2 Then Exit Function
If ListViewHeaderHandle = 0 Then ListViewHeaderHandle = Me.hWndHeader
If ListViewHeaderHandle <> 0 Then
    Dim Count As Long
    Count = Me.ColumnHeaders.Count
    If Count > 0 Then
        Dim RC As RECT
        GetClientRect hWndFilterEdit, RC
        MapWindowPoints hWndFilterEdit, ListViewHeaderHandle, RC, 2
        Dim LVC As LVCOLUMN, i As Long, CX As Long
        LVC.Mask = LVCF_WIDTH
        For i = 1 To Count
            SendMessage ListViewHandle, LVM_GETCOLUMN, i - 1, ByVal VarPtr(LVC)
            CX = CX + LVC.CX
            If CX >= RC.Left Then Exit For
        Next i
        GetFilterEditIndex = i
    End If
End If
End Function

Private Function IndexToStateImageMask(ByVal ImgIndex As Long) As Long
IndexToStateImageMask = ImgIndex * (2 ^ 12)
End Function

Private Function StateImageMaskToIndex(ByVal ImgState As Long) As Long
StateImageMaskToIndex = ImgState / (2 ^ 12)
End Function

Private Function PropIconsControl() As Object
If ListViewIconsObjectPointer <> 0 Then Set PropIconsControl = PtrToObj(ListViewIconsObjectPointer)
End Function

Private Function PropSmallIconsControl() As Object
If ListViewSmallIconsObjectPointer <> 0 Then Set PropSmallIconsControl = PtrToObj(ListViewSmallIconsObjectPointer)
End Function

Private Function PropColumnHeaderIconsControl() As Object
If ListViewColumnHeaderIconsObjectPointer <> 0 Then Set PropColumnHeaderIconsControl = PtrToObj(ListViewColumnHeaderIconsObjectPointer)
End Function

Private Function PropGroupIconsControl() As Object
If ListViewGroupIconsObjectPointer <> 0 Then Set PropGroupIconsControl = PtrToObj(ListViewGroupIconsObjectPointer)
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcLabelEdit(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcFilterEdit(hWnd, wMsg, wParam, lParam)
    Case 4
        ISubclass_Message = WindowProcHeader(hWnd, wMsg, wParam, lParam)
    Case 5
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Case 10
        ISubclass_Message = ListItemsSortingFunctionBinary(wParam, lParam)
    Case 11
        ISubclass_Message = ListItemsSortingFunctionText(wParam, lParam)
    Case 12
        ISubclass_Message = ListItemsSortingFunctionNumeric(wParam, lParam)
    Case 13
        ISubclass_Message = ListItemsSortingFunctionCurrency(wParam, lParam)
    Case 14
        ISubclass_Message = ListItemsSortingFunctionDate(wParam, lParam)
    Case 15
        ISubclass_Message = ListItemsSortingFunctionLogical(wParam, lParam)
    Case 20
        ISubclass_Message = GroupsSortingFunctionBinary(wParam, lParam, wMsg)
    Case 21
        ISubclass_Message = GroupsSortingFunctionText(wParam, lParam, wMsg)
    Case 22
        ISubclass_Message = GroupsSortingFunctionNumeric(wParam, lParam, wMsg)
    Case 23
        ISubclass_Message = GroupsSortingFunctionCurrency(wParam, lParam, wMsg)
    Case 24
        ISubclass_Message = GroupsSortingFunctionDate(wParam, lParam, wMsg)
    Case 25
        ISubclass_Message = GroupsSortingFunctionLogical(wParam, lParam, wMsg)
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
                If PropResizableColumnHeaders = False Then
                    If KeyCode = vbKeyAdd And (GetShiftStateFromMsg() And vbCtrlMask) = vbCtrlMask Then Exit Function
                End If
                If PropCheckboxes = True And KeyCode = vbKeySpace And PropVirtualMode = True Then
                    ' A virtual list view where checkboxes are displayed does not generate LVN_ITEMCHANGED upon pressing the space key.
                    Dim iItem As Long
                    iItem = SendMessage(ListViewHandle, LVM_GETNEXTITEM, -1, ByVal LVNI_ALL Or LVNI_FOCUSED)
                    If iItem > -1 Then
                        Dim ListItem As LvwListItem
                        Set ListItem = New LvwListItem
                        ListItem.FInit ObjPtr(Me), iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                        RaiseEvent ItemCheck(ListItem, Not CBool(StateImageMaskToIndex(SendMessage(ListViewHandle, LVM_GETITEMSTATE, iItem, ByVal LVIS_STATEIMAGEMASK) And LVIS_STATEIMAGEMASK) = IIL_CHECKED))
                        SendMessage ListViewHandle, LVM_UPDATE, iItem, ByVal 0&
                    End If
                End If
            ElseIf wMsg = WM_KEYUP Then
                RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
            End If
            ListViewCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If ListViewCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(ListViewCharCodeCache And &HFFFF&)
            ListViewCharCodeCache = 0
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
        Call ComCtlsSetIMEMode(hWnd, ListViewIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call ComCtlsSetIMEMode(hWnd, ListViewIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = ListViewHeaderHandle And ListViewHeaderHandle <> 0 Then
            Dim Cancel As Boolean
            Dim NMHDR As NMHEADER, HDI As HDITEM
            Select Case NM.Code
                Case HDN_ITEMDBLCLICK
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then RaiseEvent ColumnDblClick(Me.ColumnHeaders(NMHDR.iItem + 1))
                Case HDN_DIVIDERDBLCLICK
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then
                        If PropResizableColumnHeaders = True Then
                            If Me.ColumnHeaders(NMHDR.iItem + 1).Resizable = True Then
                                RaiseEvent ColumnDividerDblClick(Me.ColumnHeaders(NMHDR.iItem + 1), Cancel)
                            Else
                                Cancel = True
                            End If
                        Else
                            Cancel = True
                        End If
                        If Cancel = True Then Exit Function
                    End If
                Case HDN_BEGINTRACK
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then
                        If PropResizableColumnHeaders = True Then
                            If Me.ColumnHeaders(NMHDR.iItem + 1).Resizable = True Then
                                RaiseEvent ColumnBeforeResize(Me.ColumnHeaders(NMHDR.iItem + 1), Cancel)
                            Else
                                Cancel = True
                            End If
                        Else
                            Cancel = True
                        End If
                        If Cancel = True Then
                            WindowProcControl = 1
                            Exit Function
                        End If
                    End If
                Case HDN_ENDTRACK
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then
                        If NMHDR.lPtrHDItem <> 0 Then
                            CopyMemory HDI.Mask, ByVal NMHDR.lPtrHDItem, 4
                            If (HDI.Mask And HDI_WIDTH) = HDI_WIDTH Then
                                Dim NewWidth As Single, CX As Long
                                CopyMemory HDI.CXY, ByVal UnsignedAdd(NMHDR.lPtrHDItem, 4), 4
                                NewWidth = UserControl.ScaleX(HDI.CXY, vbPixels, vbContainerSize)
                                RaiseEvent ColumnAfterResize(Me.ColumnHeaders(NMHDR.iItem + 1), NewWidth)
                                If NewWidth > 0 Then CX = UserControl.ScaleX(NewWidth, vbContainerSize, vbPixels)
                                If HDI.CXY <> CX Then CopyMemory ByVal UnsignedAdd(NMHDR.lPtrHDItem, 4), CX, 4
                            End If
                        End If
                    End If
                Case HDN_BEGINDRAG
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then RaiseEvent ColumnBeforeDrag(Me.ColumnHeaders(NMHDR.iItem + 1))
                Case HDN_ENDDRAG
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then
                        If NMHDR.lPtrHDItem <> 0 Then
                            CopyMemory HDI.Mask, ByVal NMHDR.lPtrHDItem, 4
                            If (HDI.Mask And HDI_ORDER) = HDI_ORDER Then
                                CopyMemory HDI.iOrder, ByVal UnsignedAdd(NMHDR.lPtrHDItem, 32), 4
                                RaiseEvent ColumnAfterDrag(Me.ColumnHeaders(NMHDR.iItem + 1), HDI.iOrder + 1, Cancel)
                                If Cancel = True Then
                                    WindowProcControl = 1
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Case HDN_DROPDOWN
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then RaiseEvent ColumnDropDown(Me.ColumnHeaders(NMHDR.iItem + 1))
                Case HDN_ITEMSTATEICONCLICK
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then
                        With Me.ColumnHeaders(NMHDR.iItem + 1)
                        .Checked = Not .Checked
                        End With
                        Exit Function
                    End If
                Case HDN_FILTERCHANGE
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    If NMHDR.iItem > -1 Then RaiseEvent ColumnFilterChanged(Me.ColumnHeaders(NMHDR.iItem + 1))
                Case HDN_FILTERBTNCLICK
                    Dim NMHDFBC As NMHDFILTERBTNCLICK
                    CopyMemory NMHDFBC, ByVal lParam, LenB(NMHDFBC)
                    If NMHDFBC.iItem > -1 Then
                        Dim RaiseFilterChanged As Boolean
                        With NMHDFBC
                        RaiseEvent ColumnFilterButtonClick(Me.ColumnHeaders(.iItem + 1), RaiseFilterChanged, .RC.Left, .RC.Top, .RC.Right, .RC.Bottom)
                        End With
                        If RaiseFilterChanged = True Then WindowProcControl = 1: Exit Function
                    End If
                Case HDN_BEGINFILTEREDIT
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    ' It is necessary to overwrite iItem by HDM_GETFOCUSEDITEM as otherwise it would be always -1.
                    NMHDR.iItem = SendMessage(NMHDR.hdr.hWndFrom, HDM_GETFOCUSEDITEM, 0, ByVal 0&)
                    If NMHDR.iItem > -1 Then RaiseEvent BeforeFilterEdit(Me.ColumnHeaders(NMHDR.iItem + 1), ListViewFilterEditHandle)
                Case HDN_ENDFILTEREDIT
                    CopyMemory NMHDR, ByVal lParam, LenB(NMHDR)
                    ' It is necessary to overwrite iItem by HDM_GETFOCUSEDITEM as otherwise it would be always -1.
                    NMHDR.iItem = SendMessage(NMHDR.hdr.hWndFrom, HDM_GETFOCUSEDITEM, 0, ByVal 0&)
                    If NMHDR.iItem > -1 Then RaiseEvent AfterFilterEdit(Me.ColumnHeaders(NMHDR.iItem + 1))
                Case NM_CUSTOMDRAW
                    Dim FontHandle As Long
                    Dim NMCD As NMCUSTOMDRAW
                    CopyMemory NMCD, ByVal lParam, LenB(NMCD)
                    Select Case NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            WindowProcControl = CDRF_NOTIFYITEMDRAW
                            Exit Function
                        Case CDDS_ITEMPREPAINT
                            FontHandle = ListViewFontHandle
                            If NMCD.dwItemSpec > -1 Then
                                With Me.ColumnHeaders(NMCD.dwItemSpec + 1)
                                SetTextColor NMCD.hDC, WinColor(.ForeColor)
                                If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                End With
                            End If
                            SelectObject NMCD.hDC, FontHandle
                            WindowProcControl = CDRF_NEWFONT
                            Exit Function
                    End Select
            End Select
        ElseIf NM.hWndFrom = ListViewToolTipHandle And ListViewToolTipHandle <> 0 Then
            Static ShowSubInfoTip As Boolean
            Select Case NM.Code
                Case TTN_GETDISPINFO
                    Dim NMTTDI As NMTTDISPINFO
                    CopyMemory NMTTDI, ByVal lParam, LenB(NMTTDI)
                    If PropRightToLeft = True And PropRightToLeftLayout = False Then
                        If Not (NMTTDI.uFlags And TTF_RTLREADING) = TTF_RTLREADING Then
                            NMTTDI.uFlags = NMTTDI.uFlags Or TTF_RTLREADING
                            CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                        End If
                    End If
                    ShowSubInfoTip = False
                    If PropView = LvwViewReport Then
                        Dim LVHTI As LVHITTESTINFO, Pos As Long
                        With LVHTI
                        Pos = GetMessagePos()
                        .PT.X = Get_X_lParam(Pos)
                        .PT.Y = Get_Y_lParam(Pos)
                        ScreenToClient hWnd, .PT
                        If SendMessage(hWnd, LVM_SUBITEMHITTEST, 0, ByVal VarPtr(LVHTI)) > -1 Then
                            If (.Flags And LVHT_ONITEM) <> 0 And .iSubItem > 0 Then
                                Dim Text As String, Length As Long
                                If (SendMessage(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, ByVal 0&) And LVS_EX_LABELTIP) = LVS_EX_LABELTIP Then
                                    WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
                                    CopyMemory NMTTDI, ByVal lParam, LenB(NMTTDI)
                                    If NMTTDI.lpszText <> 0 Then Length = lstrlen(NMTTDI.lpszText)
                                    If Length > 0 Then
                                        Text = String(Length, vbNullChar)
                                        CopyMemory ByVal StrPtr(Text), ByVal NMTTDI.lpszText, Length * 2
                                    Else
                                        Text = Left$(NMTTDI.szText(), InStr(NMTTDI.szText(), vbNullChar) - 1)
                                    End If
                                End If
                                If Text = vbNullString Then
                                    If (SendMessage(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, ByVal 0&) And LVS_EX_INFOTIP) = LVS_EX_INFOTIP Then
                                        If PropVirtualMode = False Then
                                            Text = Me.ListItems(.iItem + 1).ListSubItems(.iSubItem).ToolTipText
                                        Else
                                            If (PropVirtualDisabledInfos And LvwVirtualPropertyToolTipText) = 0 Then
                                                RaiseEvent GetVirtualItem(.iItem + 1, .iSubItem, LvwVirtualPropertyToolTipText, Text)
                                            End If
                                        End If
                                        If Not Text = vbNullString Then
                                            If Len(Text) <= 80 Then
                                                Text = Left$(Text & vbNullChar, 80)
                                                CopyMemory NMTTDI.szText(0), ByVal StrPtr(Text), LenB(Text)
                                            Else
                                                Erase NMTTDI.szText()
                                            End If
                                            NMTTDI.lpszText = StrPtr(Text) ' Apparently the string address must be always set.
                                            NMTTDI.hInst = 0
                                            CopyMemory ByVal lParam, NMTTDI, LenB(NMTTDI)
                                            ShowSubInfoTip = True
                                        End If
                                    End If
                                End If
                                Exit Function
                            End If
                        End If
                        End With
                    End If
                Case TTN_SHOW
                    If (PropView <> LvwViewIcon And (SendMessage(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, ByVal 0&) And LVS_EX_LABELTIP) = 0) _
                    Or ShowSubInfoTip = True Then
                        ' To display the ToolTip in its default location, return zero.
                        WindowProcControl = 0
                        Exit Function
                    End If
            End Select
        End If
    Case UM_BUTTONDOWN
        ' The control enters a modal message loop (DragDetect) on WM_LBUTTONDOWN and WM_RBUTTONDOWN.
        ' This workaround is necessary to raise 'MouseDown' before the button was released or the mouse was moved.
        RaiseEvent MouseDown(LoWord(wParam), HiWord(wParam), UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips), UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips))
        ListViewButtonDown = LoWord(wParam)
        ListViewIsClick = True
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
                ' In case DragDetect returns 0 then the control will set focus the focus automatically.
                ' Otherwise not. So check and change focus, if needed.
                If GetFocus() <> hWnd Then SetFocusAPI hWnd
                ' See UM_BUTTONDOWN
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                ListViewButtonDown = 0
                ListViewIsClick = True
            Case WM_RBUTTONDOWN
                ' In case DragDetect returns 0 then the control will set focus the focus automatically.
                ' Otherwise not. So check and change focus, if needed.
                If GetFocus() <> hWnd Then SetFocusAPI hWnd
                ' See UM_BUTTONDOWN
            Case WM_MOUSEMOVE
                If ListViewMouseOver = False And PropMouseTrack = True Then
                    ListViewMouseOver = True
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
                ListViewButtonDown = 0
                If ListViewIsClick = True Then
                    ListViewIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If ListViewMouseOver = True Then
            ListViewMouseOver = False
            RaiseEvent MouseLeave
        End If
    Case LVM_SETHOTLIGHTCOLOR
        ' Since this is a undocumented message it is safer to support it only indirectly.
        If WindowProcControl <> 0 Then ListViewHotLightColor = SendMessage(hWnd, LVM_GETHOTLIGHTCOLOR, 0, ByVal 0&)
End Select
End Function

Private Function WindowProcLabelEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_KEYDOWN, WM_KEYUP
        ListViewCharCodeCache = ComCtlsPeekCharCode(hWnd)
    Case WM_CHAR
        If ListViewCharCodeCache <> 0 Then
            wParam = ListViewCharCodeCache
            ListViewCharCodeCache = 0
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
        Call ComCtlsSetIMEMode(hWnd, ListViewIMCHandle, PropIMEMode)
    Case WM_IME_SETCONTEXT
        If wParam <> 0 Then Call ComCtlsSetIMEMode(hWnd, ListViewIMCHandle, PropIMEMode)
    Case WM_IME_CHAR
        SendMessage hWnd, WM_CHAR, wParam, ByVal lParam
        Exit Function
End Select
WindowProcLabelEdit = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcFilterEdit(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg = WM_KILLFOCUS Then
    ' The filter edit window will be destroyed when it receives WM_KILLFOCUS.
    Call ComCtlsRemoveSubclass(hWnd)
End If
WindowProcFilterEdit = WindowProcLabelEdit(hWnd, wMsg, wParam, lParam)
If wMsg = WM_KILLFOCUS Then
    If ListViewFilterEditIndex > 0 Then RaiseEvent AfterFilterEdit(Me.ColumnHeaders(ListViewFilterEditIndex))
    ListViewFilterEditHandle = 0
    ListViewFilterEditIndex = 0
End If
End Function

Private Function WindowProcHeader(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SETFOCUS
        Call ActivateIPAO(Me)
    Case WM_KILLFOCUS
        Call DeActivateIPAO
    Case WM_LBUTTONDOWN
        Dim HDHTI1 As HDHITTESTINFO
        With HDHTI1
        .PT.X = Get_X_lParam(lParam)
        .PT.Y = Get_Y_lParam(lParam)
        If SendMessage(hWnd, HDM_HITTEST, 0, ByVal VarPtr(HDHTI1)) > -1 Then
            If (.Flags And HHT_ONFILTER) <> 0 Then
                Select Case GetFocus()
                    Case hWnd, ListViewHandle
                    Case Else
                        UCNoSetFocusFwd = True: SetFocusAPI UserControl.hWnd: UCNoSetFocusFwd = False
                End Select
            End If
        End If
        End With
    Case WM_SETCURSOR
        If LoWord(lParam) = HTCLIENT Then
            Dim hCursor As Long
            If MousePointerID(PropHeaderMousePointer) <> 0 Then
                hCursor = LoadCursor(0, MousePointerID(PropHeaderMousePointer))
            ElseIf PropHeaderMousePointer = 99 Then
                If Not PropHeaderMouseIcon Is Nothing Then hCursor = PropHeaderMouseIcon.Handle
            End If
            Dim HDHTI2 As HDHITTESTINFO, Pos As Long
            With HDHTI2
            Pos = GetMessagePos()
            .PT.X = Get_X_lParam(Pos)
            .PT.Y = Get_Y_lParam(Pos)
            ScreenToClient hWnd, .PT
            If SendMessage(hWnd, HDM_HITTEST, 0, ByVal VarPtr(HDHTI2)) > -1 Then
                If (.Flags And HHT_ONDIVIDER) <> 0 Or (.Flags And HHT_ONDIVOPEN) <> 0 Then
                    If PropResizableColumnHeaders = False Then
                        If hCursor = 0 Then hCursor = LoadCursor(0, MousePointerID(vbArrow))
                    ElseIf Me.ColumnHeaders(.iItem + 1).Resizable = False Then
                        If hCursor = 0 Then hCursor = LoadCursor(0, MousePointerID(vbArrow))
                    Else
                        hCursor = 0
                    End If
                End If
            End If
            End With
            If hCursor <> 0 Then
                SetCursor hCursor
                WindowProcHeader = 1
                Exit Function
            End If
        End If
    Case WM_SIZE
        If PropShowColumnTips = True Then Call UpdateHeaderToolTipRect(hWnd)
    Case WM_MOUSEMOVE
        If PropShowColumnTips = True Then Call CheckHeaderToolTipItem(hWnd, Get_X_lParam(lParam), Get_Y_lParam(lParam))
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = ListViewHeaderToolTipHandle And ListViewHeaderToolTipHandle <> 0 Then
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
                    Dim Text As String
                    Text = GetColumnToolTipText(hWnd, GetMessagePos())
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
    Case WM_COMMAND
        Const EN_SETFOCUS As Long = &H100, EN_KILLFOCUS As Long = &H200
        Select Case HiWord(wParam)
            Case EN_SETFOCUS
                ListViewFilterEditHandle = lParam
                ListViewFilterEditIndex = GetFilterEditIndex(lParam)
                If lParam <> 0 Then
                    If PropRightToLeft = True And PropRightToLeftLayout = False Then Call ComCtlsSetRightToLeft(lParam, WS_EX_RTLREADING)
                    Call ComCtlsSetSubclass(lParam, Me, 3)
                    Call ActivateIPAO(Me)
                End If
                If ListViewFilterEditIndex > 0 Then RaiseEvent BeforeFilterEdit(Me.ColumnHeaders(ListViewFilterEditIndex), ListViewFilterEditHandle)
            Case EN_KILLFOCUS
                ' When the user types ESC or RETURN the filter edit window sends EN_KILLFOCUS.
                ' In all other cases the filter edit window sends WM_KILLFOCUS.
                ' Thus it is necessary to handle both EN_KILLFOCUS and WM_KILLFOCUS.
                ' UM_ENDFILTEREDIT will be posted as the filter edit window is not yet destroyed.
                PostMessage hWnd, UM_ENDFILTEREDIT, ListViewFilterEditIndex, ByVal 0&
                ListViewFilterEditHandle = 0
                ListViewFilterEditIndex = 0
                If lParam <> 0 Then Call ComCtlsRemoveSubclass(lParam)
        End Select
    Case UM_ENDFILTEREDIT
        If wParam > 0 Then RaiseEvent AfterFilterEdit(Me.ColumnHeaders(wParam))
        Exit Function
End Select
WindowProcHeader = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = ListViewHandle Then
            Dim ListItem As LvwListItem
            Dim Length As Long, Cancel As Boolean
            Dim NMLV As NMLISTVIEW, NMIA As NMITEMACTIVATE, NMLVDI As NMLVDISPINFO
            Select Case NM.Code
                Case LVN_INSERTITEM
                    If ListViewListItemsControl = 0 Then
                        Dim LVI As LVITEM
                        With LVI
                        If PropAutoSelectFirstItem = True Then
                            .StateMask = LVIS_SELECTED Or LVIS_FOCUSED
                            .State = LVIS_SELECTED Or LVIS_FOCUSED
                        Else
                            .StateMask = LVIS_FOCUSED
                            .State = LVIS_FOCUSED
                        End If
                        End With
                        SendMessage ListViewHandle, LVM_SETITEMSTATE, 0, ByVal VarPtr(LVI)
                    End If
                    ListViewListItemsControl = ListViewListItemsControl + 1
                Case LVN_DELETEITEM
                    ListViewListItemsControl = ListViewListItemsControl - 1
                Case LVN_ITEMCHANGED
                    CopyMemory NMLV, ByVal lParam, LenB(NMLV)
                    With NMLV
                    If .uChanged = LVIF_STATE Then
                        If .iItem > -1 Then
                            If PropVirtualMode = False Then
                                Set ListItem = Me.ListItems(.iItem + 1)
                            Else
                                Set ListItem = New LvwListItem
                                ListItem.FInit ObjPtr(Me), .iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                            End If
                            If CBool((.uNewState And LVIS_FOCUSED) = LVIS_FOCUSED) Xor CBool((.uOldState And LVIS_FOCUSED) = LVIS_FOCUSED) Then
                                If (.uNewState And LVIS_FOCUSED) = LVIS_FOCUSED Then Call CheckItemFocus(.iItem + 1)
                            End If
                            If CBool((.uNewState And LVIS_SELECTED) = LVIS_SELECTED) Xor CBool((.uOldState And LVIS_SELECTED) = LVIS_SELECTED) Then
                                RaiseEvent ItemSelect(ListItem, CBool((.uNewState And LVIS_SELECTED) = LVIS_SELECTED))
                            End If
                            If PropVirtualMode = False Then
                                If CBool((.uNewState And &H2000&) = &H2000&) Xor CBool((.uOldState And &H2000&) = &H2000&) Then RaiseEvent ItemCheck(ListItem, CBool((.uNewState And &H2000&) = &H2000&))
                            End If
                        Else
                            ' The change has been applied to all items in the list view.
                            ' AFAIK only a virtual list view uses this alias to inform that all is deselected.
                            ' Because a virtual list view does only inform about each selected item and not for each deselected item.
                            If CBool((.uNewState And LVIS_SELECTED) = LVIS_SELECTED) Xor CBool((.uOldState And LVIS_SELECTED) = LVIS_SELECTED) Then
                                RaiseEvent ItemSelect(Nothing, CBool((.uNewState And LVIS_SELECTED) = LVIS_SELECTED))
                            End If
                        End If
                    End If
                    End With
                Case LVN_BEGINLABELEDIT, LVN_ENDLABELEDIT
                    Static LabelEditHandle As Long
                    Select Case NM.Code
                        Case LVN_BEGINLABELEDIT
                            If PropLabelEdit = LvwLabelEditManual And ListViewStartLabelEdit = False Then
                                WindowProcUserControl = 1
                            Else
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
                                    ListViewLabelInEdit = True
                                End If
                            End If
                        Case LVN_ENDLABELEDIT
                            CopyMemory NMLVDI, ByVal lParam, LenB(NMLVDI)
                            With NMLVDI.Item
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
                            ListViewLabelInEdit = False
                    End Select
                    Exit Function
                Case LVN_BEGINDRAG, LVN_BEGINRDRAG
                    CopyMemory NMLV, ByVal lParam, LenB(NMLV)
                    If NMLV.iItem > -1 Then
                        If PropVirtualMode = False Then
                            Set ListItem = Me.ListItems(NMLV.iItem + 1)
                        Else
                            Set ListItem = New LvwListItem
                            ListItem.FInit ObjPtr(Me), NMLV.iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                        End If
                        ListViewDragIndexBuffer = NMLV.iItem + 1
                        If NM.Code = LVN_BEGINDRAG Then
                            RaiseEvent ItemDrag(ListItem, vbLeftButton)
                            If PropOLEDragMode = vbOLEDragAutomatic Then Me.OLEDrag
                        ElseIf NM.Code = LVN_BEGINRDRAG Then
                            RaiseEvent ItemDrag(ListItem, vbRightButton)
                        End If
                        ListViewDragIndexBuffer = 0
                    End If
                Case LVN_COLUMNCLICK
                    CopyMemory NMLV, ByVal lParam, LenB(NMLV)
                    RaiseEvent ColumnClick(Me.ColumnHeaders(NMLV.iSubItem + 1))
                Case LVN_ITEMACTIVATE
                    CopyMemory NMIA, ByVal lParam, LenB(NMIA)
                    Dim Shift As Integer
                    If (NMIA.uKeyFlags And LVKF_SHIFT) = LVKF_SHIFT Then Shift = vbShiftMask
                    If (NMIA.uKeyFlags And LVKF_CONTROL) = LVKF_CONTROL Then Shift = Shift Or vbCtrlMask
                    If (NMIA.uKeyFlags And LVKF_ALT) = LVKF_ALT Then Shift = Shift Or vbAltMask
                    If PropVirtualMode = False Then
                        Set ListItem = Me.ListItems(NMIA.iItem + 1)
                    Else
                        Set ListItem = New LvwListItem
                        ListItem.FInit ObjPtr(Me), NMIA.iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                    End If
                    RaiseEvent ItemActivate(ListItem, NMIA.iSubItem, Shift)
                Case NM_CLICK, NM_RCLICK
                    CopyMemory NMIA, ByVal lParam, LenB(NMIA)
                    If NMIA.iItem > -1 Then
                        If PropVirtualMode = False Then
                            Set ListItem = Me.ListItems(NMIA.iItem + 1)
                        Else
                            Set ListItem = New LvwListItem
                            ListItem.FInit ObjPtr(Me), NMIA.iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                        End If
                        If NM.Code = NM_CLICK Then
                            RaiseEvent ItemClick(ListItem, vbLeftButton)
                        ElseIf NM.Code = NM_RCLICK Then
                            RaiseEvent ItemClick(ListItem, vbRightButton)
                        End If
                        If PropCheckboxes = True And PropVirtualMode = True Then
                            ' A virtual list view where checkboxes are displayed does not generate LVN_ITEMCHANGED upon clicking the checkbox.
                            Dim LVHTI As LVHITTESTINFO
                            With LVHTI
                            LSet .PT = NMIA.PTAction
                            SendMessage ListViewHandle, LVM_SUBITEMHITTEST, 0, ByVal VarPtr(LVHTI)
                            If (.Flags And LVHT_ONITEM) <> 0 And .iSubItem = 0 Then
                                If (.Flags And LVHT_ONITEMSTATEICON) <> 0 Then
                                    RaiseEvent ItemCheck(ListItem, Not CBool(StateImageMaskToIndex(SendMessage(ListViewHandle, LVM_GETITEMSTATE, NMIA.iItem, ByVal LVIS_STATEIMAGEMASK) And LVIS_STATEIMAGEMASK) = IIL_CHECKED))
                                    SendMessage ListViewHandle, LVM_UPDATE, NMIA.iItem, ByVal 0&
                                End If
                            End If
                            End With
                        End If
                    End If
                    If NMIA.iItem > -1 Or (NMIA.iItem = -1 And (PropView = LvwViewReport Or PropView = LvwViewList)) Then
                        If ListViewButtonDown <> 0 Then
                            RaiseEvent MouseUp(ListViewButtonDown, GetShiftStateFromMsg(), UserControl.ScaleX(NMIA.PTAction.X, vbPixels, vbTwips), UserControl.ScaleY(NMIA.PTAction.Y, vbPixels, vbTwips))
                            ListViewButtonDown = 0
                            ListViewIsClick = False
                            RaiseEvent Click
                        End If
                    End If
                Case NM_DBLCLK, NM_RDBLCLK
                    CopyMemory NMIA, ByVal lParam, LenB(NMIA)
                    If NMIA.iItem > -1 Then
                        If PropVirtualMode = False Then
                            Set ListItem = Me.ListItems(NMIA.iItem + 1)
                        Else
                            Set ListItem = New LvwListItem
                            ListItem.FInit ObjPtr(Me), NMIA.iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                        End If
                        If NM.Code = NM_DBLCLK Then
                            RaiseEvent ItemDblClick(ListItem, vbLeftButton)
                        ElseIf NM.Code = NM_RDBLCLK Then
                            RaiseEvent ItemDblClick(ListItem, vbRightButton)
                        End If
                    End If
                    RaiseEvent DblClick
                Case NM_CUSTOMDRAW
                    Dim FontHandle As Long, Bold As Boolean, ForeColor As OLE_COLOR
                    Dim NMLVCD As NMLVCUSTOMDRAW
                    CopyMemory NMLVCD, ByVal lParam, LenB(NMLVCD)
                    Select Case NMLVCD.NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            WindowProcUserControl = CDRF_NOTIFYITEMDRAW
                            Exit Function
                        Case CDDS_ITEMPREPAINT
                            FontHandle = ListViewFontHandle
                            If PropVirtualMode = True Then
                                If NMLVCD.NMCD.dwItemSpec > -1 And NMLVCD.NMCD.dwItemSpec <= PropVirtualItemCount Then
                                    If NMLVCD.iSubItem = 0 Then
                                        If (NMLVCD.NMCD.uItemState And CDIS_HOT) = 0 Or PropHotTracking = False Then
                                            If (PropVirtualDisabledInfos And LvwVirtualPropertyBold) = 0 Then
                                                RaiseEvent GetVirtualItem(NMLVCD.NMCD.dwItemSpec + 1, NMLVCD.iSubItem, LvwVirtualPropertyBold, Bold)
                                            End If
                                            If Bold = True Then FontHandle = ListViewBoldFontHandle
                                            ForeColor = PropForeColor
                                            If (PropVirtualDisabledInfos And LvwVirtualPropertyForeColor) = 0 Then
                                                RaiseEvent GetVirtualItem(NMLVCD.NMCD.dwItemSpec + 1, NMLVCD.iSubItem, LvwVirtualPropertyForeColor, ForeColor)
                                            End If
                                            NMLVCD.ClrText = WinColor(ForeColor)
                                        Else
                                            If (PropVirtualDisabledInfos And LvwVirtualPropertyBold) = 0 Then
                                                RaiseEvent GetVirtualItem(NMLVCD.NMCD.dwItemSpec + 1, NMLVCD.iSubItem, LvwVirtualPropertyBold, Bold)
                                            End If
                                            If PropUnderlineHot = True Then
                                                If Bold = True Then
                                                    FontHandle = ListViewBoldUnderlineFontHandle
                                                Else
                                                    FontHandle = ListViewUnderlineFontHandle
                                                End If
                                            Else
                                                If Bold = True Then FontHandle = ListViewBoldFontHandle
                                            End If
                                            If PropHighlightHot = True Then
                                                If ListViewHotLightColor = CLR_DEFAULT Then
                                                    NMLVCD.ClrText = GetSysColor(COLOR_HOTLIGHT)
                                                Else
                                                    NMLVCD.ClrText = ListViewHotLightColor
                                                End If
                                            Else
                                                ForeColor = PropForeColor
                                                If (PropVirtualDisabledInfos And LvwVirtualPropertyForeColor) = 0 Then
                                                    RaiseEvent GetVirtualItem(NMLVCD.NMCD.dwItemSpec + 1, NMLVCD.iSubItem, LvwVirtualPropertyForeColor, ForeColor)
                                                End If
                                                NMLVCD.ClrText = WinColor(ForeColor)
                                            End If
                                        End If
                                        Set ListItem = New LvwListItem
                                        ListItem.FInit ObjPtr(Me), NMLVCD.NMCD.dwItemSpec + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                                        RaiseEvent ItemBkColor(ListItem, NMLVCD.ClrTextBk)
                                    End If
                                End If
                            ElseIf NMLVCD.NMCD.lItemlParam <> 0 Then
                                Set ListItem = PtrToObj(NMLVCD.NMCD.lItemlParam)
                                With ListItem
                                If (NMLVCD.NMCD.uItemState And CDIS_HOT) = 0 Or PropHotTracking = False Then
                                    If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                    NMLVCD.ClrText = WinColor(.ForeColor)
                                Else
                                    If PropUnderlineHot = True Then
                                        If .Bold = True Then
                                            FontHandle = ListViewBoldUnderlineFontHandle
                                        Else
                                            FontHandle = ListViewUnderlineFontHandle
                                        End If
                                    Else
                                        If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                    End If
                                    If PropHighlightHot = True Then
                                        If ListViewHotLightColor = CLR_DEFAULT Then
                                            NMLVCD.ClrText = GetSysColor(COLOR_HOTLIGHT)
                                        Else
                                            NMLVCD.ClrText = ListViewHotLightColor
                                        End If
                                    Else
                                        NMLVCD.ClrText = WinColor(.ForeColor)
                                    End If
                                End If
                                RaiseEvent ItemBkColor(ListItem, NMLVCD.ClrTextBk)
                                End With
                            End If
                            SelectObject NMLVCD.NMCD.hDC, FontHandle
                            CopyMemory ByVal lParam, NMLVCD, LenB(NMLVCD)
                            WindowProcUserControl = CDRF_NEWFONT Or CDRF_NOTIFYSUBITEMDRAW
                            Exit Function
                        Case (CDDS_ITEMPREPAINT Or CDDS_SUBITEM)
                            FontHandle = ListViewFontHandle
                            If PropVirtualMode = True Then
                                If NMLVCD.NMCD.dwItemSpec > -1 And NMLVCD.NMCD.dwItemSpec <= PropVirtualItemCount Then
                                    Dim SubItemCount As Long
                                    SubItemCount = Me.ColumnHeaders.Count - 1 ' Deduct 1 for SubItem 0
                                    If NMLVCD.iSubItem >= 0 And NMLVCD.iSubItem <= SubItemCount Then
                                        If (NMLVCD.NMCD.uItemState And CDIS_HOT) = 0 Or PropHotTracking = False Then
                                            If (PropVirtualDisabledInfos And LvwVirtualPropertyBold) = 0 Then
                                                RaiseEvent GetVirtualItem(NMLVCD.NMCD.dwItemSpec + 1, NMLVCD.iSubItem, LvwVirtualPropertyBold, Bold)
                                            End If
                                            If Bold = True Then FontHandle = ListViewBoldFontHandle
                                            ForeColor = PropForeColor
                                            If (PropVirtualDisabledInfos And LvwVirtualPropertyForeColor) = 0 Then
                                                RaiseEvent GetVirtualItem(NMLVCD.NMCD.dwItemSpec + 1, NMLVCD.iSubItem, LvwVirtualPropertyForeColor, ForeColor)
                                            End If
                                            NMLVCD.ClrText = WinColor(ForeColor)
                                        Else
                                            If (PropVirtualDisabledInfos And LvwVirtualPropertyBold) = 0 Then
                                                RaiseEvent GetVirtualItem(NMLVCD.NMCD.dwItemSpec + 1, NMLVCD.iSubItem, LvwVirtualPropertyBold, Bold)
                                            End If
                                            If PropUnderlineHot = True And PropView = LvwViewReport Then
                                                If Bold = True Then
                                                    FontHandle = ListViewBoldUnderlineFontHandle
                                                Else
                                                    FontHandle = ListViewUnderlineFontHandle
                                                End If
                                            Else
                                                If Bold = True Then FontHandle = ListViewBoldFontHandle
                                            End If
                                            If PropHighlightHot = True And PropView = LvwViewReport Then
                                                If ListViewHotLightColor = CLR_DEFAULT Then
                                                    NMLVCD.ClrText = GetSysColor(COLOR_HOTLIGHT)
                                                Else
                                                    NMLVCD.ClrText = ListViewHotLightColor
                                                End If
                                            Else
                                                ForeColor = PropForeColor
                                                If (PropVirtualDisabledInfos And LvwVirtualPropertyForeColor) = 0 Then
                                                    RaiseEvent GetVirtualItem(NMLVCD.NMCD.dwItemSpec + 1, NMLVCD.iSubItem, LvwVirtualPropertyForeColor, ForeColor)
                                                End If
                                                NMLVCD.ClrText = WinColor(ForeColor)
                                            End If
                                        End If
                                    End If
                                    If NMLVCD.iSubItem = 0 Then
                                        Set ListItem = New LvwListItem
                                        ListItem.FInit ObjPtr(Me), NMLVCD.NMCD.dwItemSpec + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                                        RaiseEvent ItemBkColor(ListItem, NMLVCD.ClrTextBk)
                                    End If
                                End If
                            ElseIf NMLVCD.NMCD.lItemlParam <> 0 Then
                                Set ListItem = PtrToObj(NMLVCD.NMCD.lItemlParam)
                                With ListItem
                                If NMLVCD.iSubItem > 0 Then
                                    If .FListSubItemsCount > 0 Then
                                        If NMLVCD.iSubItem <= .FListSubItemsCount Then
                                            If (NMLVCD.NMCD.uItemState And CDIS_HOT) = 0 Or PropHotTracking = False Then
                                                If .FListSubItemProp(NMLVCD.iSubItem, 6) = True Then FontHandle = ListViewBoldFontHandle
                                                If .FListSubItemProp(NMLVCD.iSubItem, 7) = -1 Then
                                                    NMLVCD.ClrText = WinColor(PropForeColor)
                                                Else
                                                    NMLVCD.ClrText = WinColor(.FListSubItemProp(NMLVCD.iSubItem, 7))
                                                End If
                                            Else
                                                If PropUnderlineHot = True And PropView = LvwViewReport Then
                                                    If .FListSubItemProp(NMLVCD.iSubItem, 6) = True Then
                                                        FontHandle = ListViewBoldUnderlineFontHandle
                                                    Else
                                                        FontHandle = ListViewUnderlineFontHandle
                                                    End If
                                                Else
                                                    If .FListSubItemProp(NMLVCD.iSubItem, 6) = True Then FontHandle = ListViewBoldFontHandle
                                                End If
                                                If PropHighlightHot = True And PropView = LvwViewReport Then
                                                    If ListViewHotLightColor = CLR_DEFAULT Then
                                                        NMLVCD.ClrText = GetSysColor(COLOR_HOTLIGHT)
                                                    Else
                                                        NMLVCD.ClrText = ListViewHotLightColor
                                                    End If
                                                Else
                                                    If .FListSubItemProp(NMLVCD.iSubItem, 7) = -1 Then
                                                        NMLVCD.ClrText = WinColor(PropForeColor)
                                                    Else
                                                        NMLVCD.ClrText = WinColor(.FListSubItemProp(NMLVCD.iSubItem, 7))
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If (NMLVCD.NMCD.uItemState And CDIS_HOT) = 0 Or PropHotTracking = False Then
                                        If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                        NMLVCD.ClrText = WinColor(.ForeColor)
                                    Else
                                        If PropUnderlineHot = True Then
                                            If .Bold = True Then
                                                FontHandle = ListViewBoldUnderlineFontHandle
                                            Else
                                                FontHandle = ListViewUnderlineFontHandle
                                            End If
                                        Else
                                            If .Bold = True Then FontHandle = ListViewBoldFontHandle
                                        End If
                                        If PropHighlightHot = True Then
                                            If ListViewHotLightColor = CLR_DEFAULT Then
                                                NMLVCD.ClrText = GetSysColor(COLOR_HOTLIGHT)
                                            Else
                                                NMLVCD.ClrText = ListViewHotLightColor
                                            End If
                                        Else
                                            NMLVCD.ClrText = WinColor(.ForeColor)
                                        End If
                                    End If
                                    RaiseEvent ItemBkColor(ListItem, NMLVCD.ClrTextBk)
                                End If
                                End With
                            End If
                            SelectObject NMLVCD.NMCD.hDC, FontHandle
                            CopyMemory ByVal lParam, NMLVCD, LenB(NMLVCD)
                            WindowProcUserControl = CDRF_NEWFONT
                            Exit Function
                    End Select
                Case LVN_GETINFOTIP
                    Dim NMLVGIT As NMLVGETINFOTIP
                    CopyMemory NMLVGIT, ByVal lParam, LenB(NMLVGIT)
                    With NMLVGIT
                    If .iItem > -1 And .pszText <> 0 Then
                        If .iSubItem = 0 Then
                            Dim ToolTipText As String
                            If .dwFlags = LVGIT_UNFOLDED Or (PropView <> LvwViewIcon And (SendMessage(ListViewHandle, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, ByVal 0&) And LVS_EX_LABELTIP) = 0) Then
                                If PropVirtualMode = False Then
                                    ToolTipText = Me.ListItems(.iItem + 1).ToolTipText
                                Else
                                    If (PropVirtualDisabledInfos And LvwVirtualPropertyToolTipText) = 0 Then
                                        RaiseEvent GetVirtualItem(.iItem + 1, .iSubItem, LvwVirtualPropertyToolTipText, ToolTipText)
                                    End If
                                End If
                                If Not ToolTipText = vbNullString Then
                                    ToolTipText = Left$(ToolTipText, .cchTextMax - 1) & vbNullChar
                                    CopyMemory ByVal .pszText, ByVal StrPtr(ToolTipText), LenB(ToolTipText)
                                Else
                                    CopyMemory ByVal .pszText, 0&, 4
                                End If
                            End If
                        Else
                            ' Not supported.
                        End If
                    End If
                    End With
                Case LVN_GETDISPINFO
                    CopyMemory NMLVDI, ByVal lParam, LenB(NMLVDI)
                    With NMLVDI.Item
                    If .iItem > -1 Then
                        Dim CallbackText As String
                        If PropVirtualMode = True Then
                            If (.Mask And LVIF_TEXT) = LVIF_TEXT Then
                                If (PropVirtualDisabledInfos And LvwVirtualPropertyText) = 0 Then
                                    RaiseEvent GetVirtualItem(.iItem + 1, .iSubItem, LvwVirtualPropertyText, CallbackText)
                                End If
                                If Not CallbackText = vbNullString Then
                                    CallbackText = Left$(CallbackText, .cchTextMax - 1) & vbNullChar
                                    CopyMemory ByVal .pszText, ByVal StrPtr(CallbackText), LenB(CallbackText)
                                Else
                                    CopyMemory ByVal .pszText, 0&, 4
                                End If
                            End If
                            If (.Mask And LVIF_IMAGE) = LVIF_IMAGE Then
                                Dim Icon As Variant
                                If (PropVirtualDisabledInfos And LvwVirtualPropertyIcon) = 0 Then
                                    RaiseEvent GetVirtualItem(.iItem + 1, .iSubItem, LvwVirtualPropertyIcon, Icon)
                                End If
                                If IsEmpty(Icon) Then
                                    .iImage = -1
                                Else
                                    Dim IconIndex As Long
                                    If PropView = LvwViewIcon Then
                                        Call ComCtlsImlListImageIndex(Me, Me.Icons, Icon, IconIndex)
                                    Else
                                        Call ComCtlsImlListImageIndex(Me, Me.SmallIcons, Icon, IconIndex)
                                    End If
                                    .iImage = IconIndex - 1
                                End If
                            End If
                            If (.Mask And LVIF_INDENT) = LVIF_INDENT Then
                                Dim Indentation As Long
                                If (PropVirtualDisabledInfos And LvwVirtualPropertyIndentation) = 0 Then
                                    RaiseEvent GetVirtualItem(.iItem + 1, .iSubItem, LvwVirtualPropertyIndentation, Indentation)
                                End If
                                .iIndent = Indentation
                            End If
                            If (.Mask And LVIF_STATE) = LVIF_STATE Then
                                If (.StateMask And LVIS_STATEIMAGEMASK) = LVIS_STATEIMAGEMASK And .iSubItem = 0 Then
                                    Dim Checked As Boolean
                                    If (PropVirtualDisabledInfos And LvwVirtualPropertyChecked) = 0 Then
                                        RaiseEvent GetVirtualItem(.iItem + 1, .iSubItem, LvwVirtualPropertyChecked, Checked)
                                    End If
                                    If Checked = True Then
                                        .State = .State Or IndexToStateImageMask(IIL_CHECKED)
                                    Else
                                        .State = .State Or IndexToStateImageMask(IIL_UNCHECKED)
                                    End If
                                End If
                            End If
                            CopyMemory ByVal lParam, NMLVDI, LenB(NMLVDI)
                        ElseIf .lParam <> 0 Then
                            Set ListItem = PtrToObj(.lParam)
                            If .iSubItem = 0 Then
                                If (.Mask And LVIF_TEXT) = LVIF_TEXT Then
                                    CallbackText = ListItem.Text
                                    If Not CallbackText = vbNullString Then
                                        CallbackText = Left$(CallbackText, .cchTextMax - 1) & vbNullChar
                                        CopyMemory ByVal .pszText, ByVal StrPtr(CallbackText), LenB(CallbackText)
                                    Else
                                        CopyMemory ByVal .pszText, 0&, 4
                                    End If
                                End If
                                If (.Mask And LVIF_IMAGE) = LVIF_IMAGE Then
                                    Select Case PropView
                                        Case LvwViewIcon, LvwViewTile
                                            .iImage = ListItem.IconIndex - 1
                                        Case LvwViewSmallIcon, LvwViewList, LvwViewReport
                                            .iImage = ListItem.SmallIconIndex - 1
                                    End Select
                                End If
                            Else
                                If (.Mask And LVIF_TEXT) = LVIF_TEXT Then
                                    If .iSubItem <= ListItem.FListSubItemsCount Then CallbackText = ListItem.FListSubItemProp(.iSubItem, 3)
                                    If Not CallbackText = vbNullString Then
                                        CallbackText = Left$(CallbackText, .cchTextMax - 1) & vbNullChar
                                        CopyMemory ByVal .pszText, ByVal StrPtr(CallbackText), LenB(CallbackText)
                                    Else
                                        CopyMemory ByVal .pszText, 0&, 4
                                    End If
                                End If
                                If (.Mask And LVIF_IMAGE) = LVIF_IMAGE Then
                                    If .iSubItem <= ListItem.FListSubItemsCount Then .iImage = ListItem.FListSubItemProp(.iSubItem, 5) - 1
                                End If
                            End If
                            CopyMemory ByVal lParam, NMLVDI, LenB(NMLVDI)
                        End If
                    End If
                    End With
                Case LVN_SETDISPINFO
                    CopyMemory NMLVDI, ByVal lParam, LenB(NMLVDI)
                    With NMLVDI.Item
                    If .iItem > -1 Then
                        If PropVirtualMode = True Then
                            ' Ignore as LVN_ENDLABELEDIT is sufficient to update the text property in a virtual list view.
                        ElseIf .lParam <> 0 Then
                            Set ListItem = PtrToObj(.lParam)
                            If .iSubItem = 0 Then
                                If (.Mask And LVIF_TEXT) = LVIF_TEXT Then
                                    Dim SetText As String
                                    If .pszText <> 0 Then Length = lstrlen(.pszText)
                                    If Length > 0 Then
                                        SetText = String(Length, vbNullChar)
                                        CopyMemory ByVal StrPtr(SetText), ByVal .pszText, Length * 2
                                    End If
                                    With ListItem
                                    .FInit ObjPtr(Me), .Index, .Key, NMLVDI.Item.lParam, SetText, .Icon, .IconIndex, .SmallIcon, .SmallIconIndex
                                    End With
                                End If
                            Else
                                ' Not supported.
                            End If
                        End If
                    End If
                    End With
                Case LVN_ODFINDITEM
                    Dim NMLVFI As NMLVFINDITEM
                    CopyMemory NMLVFI, ByVal lParam, LenB(NMLVFI)
                    If (NMLVFI.LVFI.Flags And LVFI_STRING) = LVFI_STRING Then
                        Dim SearchText As String, FoundIndex As Long
                        If NMLVFI.LVFI.psz <> 0 Then Length = lstrlen(NMLVFI.LVFI.psz)
                        If Length > 0 Then
                            SearchText = String(Length, vbNullChar)
                            CopyMemory ByVal StrPtr(SearchText), ByVal NMLVFI.LVFI.psz, Length * 2
                        End If
                        RaiseEvent FindVirtualItem(NMLVFI.iStart + 1, SearchText, CBool((NMLVFI.LVFI.Flags And LVFI_PARTIAL) = LVFI_PARTIAL), CBool((NMLVFI.LVFI.Flags And LVFI_WRAP) = LVFI_WRAP), FoundIndex)
                        If FoundIndex >= 0 Then
                            WindowProcUserControl = FoundIndex - 1
                        Else
                            WindowProcUserControl = -1
                        End If
                        Exit Function
                    End If
                Case LVN_ODCACHEHINT
                    Dim NMLVCH As NMLVCACHEHINT
                    CopyMemory NMLVCH, ByVal lParam, LenB(NMLVCH)
                    RaiseEvent CacheVirtualItems(NMLVCH.iFrom + 1, NMLVCH.iTo + 1)
                Case LVN_ODSTATECHANGED
                    Dim NMLVSC As NMLVODSTATECHANGE, iItem As Long
                    CopyMemory NMLVSC, ByVal lParam, LenB(NMLVSC)
                    With NMLVSC
                    If CBool((.uNewState And LVIS_SELECTED) = LVIS_SELECTED) Xor CBool((.uOldState And LVIS_SELECTED) = LVIS_SELECTED) Then
                        Set ListItem = New LvwListItem
                        For iItem = .iFrom To .iTo
                            ListItem.FInit ObjPtr(Me), iItem + 1, vbNullString, 0, vbNullString, 0, 0, 0, 0
                            RaiseEvent ItemSelect(ListItem, CBool((.uNewState And LVIS_SELECTED) = LVIS_SELECTED))
                        Next iItem
                    End If
                    End With
                Case LVN_HOTTRACK
                    If PropHotTracking = True And PropView = LvwViewReport Then
                        ' Solve redrawing issue when the cursor moves horizontally. (e.g. from one subitem to another)
                        CopyMemory NMLV, ByVal lParam, LenB(NMLV)
                        With NMLV
                        If .iItem > -1 Then
                            If .iItem = ListViewHotTrackItem And .iSubItem <> ListViewHotTrackSubItem Then SendMessage ListViewHandle, LVM_UPDATE, .iItem, ByVal 0&
                        End If
                        ListViewHotTrackItem = .iItem
                        ListViewHotTrackSubItem = .iSubItem
                        End With
                    End If
                    WindowProcUserControl = 0
                    Exit Function
                Case LVN_GETEMPTYMARKUP
                    Dim Text As String, Center As Boolean
                    RaiseEvent GetEmptyMarkup(Text, Center)
                    If Not Text = vbNullString Then
                        Dim NMLVEMU As NMLVEMPTYMARKUP
                        CopyMemory NMLVEMU, ByVal lParam, LenB(NMLVEMU)
                        If PropRightToLeft = True And PropRightToLeftLayout = False Then Text = ChrW(&H202B) & Text ' Right-to-left Embedding (RLE)
                        Text = Left$(Text & vbNullChar, L_MAX_URL_LENGTH)
                        CopyMemory NMLVEMU.szMarkup(0), ByVal StrPtr(Text), LenB(Text)
                        If Center = True Then NMLVEMU.dwFlags = EMF_CENTERED Else NMLVEMU.dwFlags = 0
                        CopyMemory ByVal lParam, NMLVEMU, LenB(NMLVEMU)
                        WindowProcUserControl = 1
                    Else
                        WindowProcUserControl = 0
                    End If
                    Exit Function
                Case LVN_MARQUEEBEGIN
                    RaiseEvent BeginMarqueeSelection(Cancel)
                    If Cancel = True Then
                        WindowProcUserControl = 1
                    Else
                        WindowProcUserControl = 0
                    End If
                    Exit Function
                Case LVN_COLUMNOVERFLOWCLICK
                    CopyMemory NMLV, ByVal lParam, LenB(NMLV)
                    RaiseEvent ColumnChevronPushed(Me.ColumnHeaders(NMLV.iSubItem + 1))
                Case LVN_BEGINSCROLL, LVN_ENDSCROLL
                    Dim NMLVS As NMLVSCROLL
                    CopyMemory NMLVS, ByVal lParam, LenB(NMLVS)
                    If NM.Code = LVN_BEGINSCROLL Then
                        RaiseEvent BeforeScroll(UserControl.ScaleX(NMLVS.DX, vbPixels, vbContainerPosition), UserControl.ScaleY(NMLVS.DY, vbPixels, vbContainerPosition))
                    ElseIf NM.Code = LVN_ENDSCROLL Then
                        RaiseEvent AfterScroll(UserControl.ScaleX(NMLVS.DX, vbPixels, vbContainerPosition), UserControl.ScaleY(NMLVS.DY, vbPixels, vbContainerPosition))
                    End If
                Case LVN_LINKCLICK
                    Dim NMLVL As NMLVLINK
                    CopyMemory NMLVL, ByVal lParam, LenB(NMLVL)
                    With NMLVL
                    If .iGroupId > 0 Then RaiseEvent GroupLinkClick(GetGroupFromID(.iGroupId))
                    End With
                Case LVN_GROUPCHANGED
                    Dim NMLVG As NMLVGROUP
                    CopyMemory NMLVG, ByVal lParam, LenB(NMLVG)
                    With NMLVG
                    If .iGroupId > 0 Then
                        If CBool((.uNewState And LVGS_COLLAPSED) = LVGS_COLLAPSED) Xor CBool((.uOldState And LVGS_COLLAPSED) = LVGS_COLLAPSED) Then
                            RaiseEvent GroupCollapsedChanged(GetGroupFromID(.iGroupId))
                        End If
                        If CBool((.uNewState And (LVGS_SELECTED Or LVGS_FOCUSED)) = (LVGS_SELECTED Or LVGS_FOCUSED)) Xor CBool((.uOldState And (LVGS_SELECTED Or LVGS_FOCUSED)) = (LVGS_SELECTED Or LVGS_FOCUSED)) Then
                            RaiseEvent GroupSelectedChanged(GetGroupFromID(.iGroupId))
                        End If
                    End If
                    End With
            End Select
        End If
    Case WM_CONTEXTMENU
        If wParam = ListViewHandle Then
            Dim P As POINTAPI
            P.X = Get_X_lParam(lParam)
            P.Y = Get_Y_lParam(lParam)
            If P.X = -1 And P.Y = -1 Then
                ' If the user types SHIFT + F10 then the X and Y coordinates are -1.
                RaiseEvent ContextMenu(-1, -1)
            Else
                ScreenToClient ListViewHandle, P
                RaiseEvent ContextMenu(UserControl.ScaleX(P.X, vbPixels, vbContainerPosition), UserControl.ScaleY(P.Y, vbPixels, vbContainerPosition))
            End If
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
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI ListViewHandle
End Function
