VERSION 5.00
Begin VB.UserControl ToolBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ControlContainer=   -1  'True
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "ToolBar.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ToolBar.ctx":005B
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ToolBar"
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
Private TbrStyleStandard, TbrStyleFlat
Private TbrTextAlignBottom, TbrTextAlignRight
Private TbrOrientationHorizontal, TbrOrientationVertical
Private TbrButtonStyleDefault, TbrButtonStyleCheck, TbrButtonStyleCheckGroup, TbrButtonStyleSeparator, TbrButtonStyleDropDown, TbrButtonStyleWholeDropDown
Private TbrButtonValueUnpressed, TbrButtonValuePressed
#End If
Public Enum TbrStyleConstants
TbrStyleStandard = 0
TbrStyleFlat = 1
End Enum
Public Enum TbrTextAlignConstants
TbrTextAlignBottom = 0
TbrTextAlignRight = 1
End Enum
Public Enum TbrOrientationConstants
TbrOrientationHorizontal = 0
TbrOrientationVertical = 1
End Enum
Public Enum TbrButtonStyleConstants
TbrButtonStyleDefault = 0
TbrButtonStyleCheck = 1
TbrButtonStyleCheckGroup = 2
TbrButtonStyleSeparator = 3
TbrButtonStyleDropDown = 4
TbrButtonStyleWholeDropDown = 5
End Enum
Public Enum TbrButtonValueConstants
TbrButtonValueUnpressed = 0
TbrButtonValuePressed = 1
End Enum
Private Type KEYBDINPUT
wVKey As Integer
wScan As Integer
dwFlags As Long
Time As Long
dwExtraInfo As LongPtr
dwPadding As Currency ' 8 extra bytes for MOUSEINPUT.
End Type
Private Type GENERALINPUT
dwType As Long
KEYBDI As KEYBDINPUT
End Type
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
Private Type TBBUTTON
iBitmap As Long
IDCommand As Long
fsState As Byte
fsStyle As Byte
bReserved1 As Byte
bReserved2 As Byte
dwData As LongPtr
iString As LongPtr
End Type
Private Type TBBUTTONINFO
cbSize As Long
dwMask As Long
IDCommand As Long
iImage As Long
fsState As Byte
fsStyle As Byte
CX As Integer
lParam As LongPtr
pszText As LongPtr
cchText As Long
End Type
Private Type TBSAVEPARAMS
hKey As LongPtr
pszSubKey As LongPtr
pszValueName As LongPtr
End Type
Private Type TPMPARAMS
cbSize As Long
RCExclude As RECT
End Type
Private Type TBINSERTMARK
iButton As Long
dwFlags As Long
End Type
Private Type MENUINFO
cbSize As Long
fMask As Long
dwStyle As Long
CYMax As Long
hBrBack As LongPtr
dwContextHelpID As Long
dwMenuData As LongPtr
End Type
Private Type MENUITEMINFO
cbSize As Long
fMask As Long
fType As Long
fState As Long
wID As Long
hSubMenu As LongPtr
hBmpChecked As LongPtr
hBmpUnchecked As LongPtr
dwItemData As LongPtr
dwTypeData As LongPtr
cch As Long
hBmpItem As LongPtr
End Type
Private Type PAINTSTRUCT
hDC As LongPtr
fErase As Long
RCPaint As RECT
fRestore As Long
fIncUpdate As Long
RGBReserved(0 To 31) As Byte
End Type
Private Type NMHDR
hWndFrom As LongPtr
IDFrom As LongPtr
Code As Long
End Type
Private Const CDDS_PREPAINT As Long = &H1
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM + 1)
Private Const CDIS_HOT As Long = &H40
Private Const CDIS_MARKED As Long = &H80
Private Const CDRF_DODEFAULT As Long = &H0
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20
Private Const TBCDRF_HILITEHOTTRACK As Long = &H20000
Private Const TBCDRF_BLENDICON As Long = &H200000
Private Const TBCDRF_USECDCOLORS As Long = &H800000
Private Type NMCUSTOMDRAW
hdr As NMHDR
dwDrawStage As Long
hDC As LongPtr
RC As RECT
dwItemSpec As LongPtr
uItemState As Long
lItemlParam As LongPtr
End Type
Private Type NMTBCUSTOMDRAW
NMCD As NMCUSTOMDRAW
hBrMonoDither As LongPtr
hBrLines As LongPtr
hPenLines As LongPtr
ClrText As Long
ClrMark As Long
ClrTextHighlight As Long
ClrBtnFace As Long
ClrBtnHighlight As Long
ClrHighlightHotTrack As Long
RCText As RECT
nStringBkMode As Long
nHLStringBkMode As Long
End Type
Private Type NMTOOLBAR
hdr As NMHDR
iItem As Long
TBB As TBBUTTON
cchText As Long
pszText As LongPtr
End Type
Private Type NMTBDISPINFO
hdr As NMHDR
dwMask As Long
IDCommand As Long
lParam As LongPtr
iImage As Long
pszText As LongPtr
cchText As Long
End Type
Private Type NMTBGETINFOTIP
hdr As NMHDR
pszText As LongPtr
cchTextMax As Long
iItem As Long
lParam As LongPtr
End Type
Private Type NMTOOLTIPSCREATED
hdr As NMHDR
hWndToolTips As LongPtr
End Type
Private Type NMTBHOTITEM
hdr As NMHDR
IDOld As Long
IDNew As Long
dwFlags As Long
End Type
Private Type NMTBSAVE
hdr As NMHDR
lpData As LongPtr
lpCurrent As LongPtr
cbData As Long
iItem As Long
cButtons As Long
TBB As TBBUTTON
End Type
Private Type NMTBRESTORE
hdr As NMHDR
lpData As LongPtr
lpCurrent As LongPtr
cbData As Long
iItem As Long
cButtons As Long
cbBytesPerRecord As Long
TBB As TBBUTTON
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event BeginCustomization()
Attribute BeginCustomization.VB_Description = "Occurs at the begin of a customization."
Public Event InitCustomizationDialog(ByVal hDlg As Long, ByRef HideHelpButton As Boolean)
Attribute InitCustomizationDialog.VB_Description = "Occurs when the customization dialog has finished initializing."
Public Event CustomizationChange()
Attribute CustomizationChange.VB_Description = "Occurs whenever the control was customized."
Public Event ResetCustomizations(ByRef CloseDialog As Boolean)
Attribute ResetCustomizations.VB_Description = "Occurs when the user pressed the reset button in the customization dialog."
Public Event CustomizationHelp()
Attribute CustomizationHelp.VB_Description = "Occurs when the user presses the help button in the customization dialog."
Public Event EndCustomization()
Attribute EndCustomization.VB_Description = "Occurs at the end of a customization."
Public Event ButtonClick(ByVal Button As TbrButton)
Attribute ButtonClick.VB_Description = "Occurs when the user clicks on a button."
Public Event ButtonDrag(ByVal Button As TbrButton, ByVal MouseButton As Integer)
Attribute ButtonDrag.VB_Description = "Occurs when a button initiate a drag-and-drop operation."
Public Event ButtonHotChanged(ByVal Button As TbrButton, ByVal Hot As Boolean)
Attribute ButtonHotChanged.VB_Description = "Occurrs when the hot state of a button changes."
Public Event ButtonDropDown(ByVal Button As TbrButton)
Attribute ButtonDropDown.VB_Description = "Occurs when the user clicks the dropdown arrow on a button with a button style set to dropdown."
Public Event ButtonMenuClick(ByVal ButtonMenu As TbrButtonMenu)
Attribute ButtonMenuClick.VB_Description = "Occurs when the user selects an item from a button dropdown menu."
Public Event ButtonMouseEnter(ByVal Button As TbrButton)
Attribute ButtonMouseEnter.VB_Description = "Occurs when the user moves the mouse into a button."
Public Event ButtonMouseLeave(ByVal Button As TbrButton)
Attribute ButtonMouseLeave.VB_Description = "Occurs when the user moves the mouse out of a button."
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
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function VkKeyScan Lib "user32" Alias "VkKeyScanW" (ByVal cChar As Integer) As Integer
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPaint As PAINTSTRUCT) As LongPtr
Private Declare PtrSafe Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare PtrSafe Function WindowFromDC Lib "user32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, ByRef lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hDC As LongPtr, ByVal hBrush As LongPtr, ByVal lpDrawStateProc As LongPtr, ByVal lData As LongPtr, ByVal wData As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal fFlags As Long) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Declare PtrSafe Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Private Declare PtrSafe Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As LongPtr, ByRef CX As Long, ByRef CY As Long) As Long
Private Declare PtrSafe Function CreatePopupMenu Lib "user32" () As LongPtr
Private Declare PtrSafe Function DestroyMenu Lib "user32" (ByVal hMenu As LongPtr) As Long
Private Declare PtrSafe Function InsertMenuItem Lib "user32" Alias "InsertMenuItemW" (ByVal hMenu As LongPtr, ByVal uItem As Long, ByVal fByPosition As Long, ByRef lpMII As MENUITEMINFO) As Long
Private Declare PtrSafe Function SetMenuInfo Lib "user32" (ByVal hMenu As LongPtr, ByRef MI As MENUINFO) As Long
Private Declare PtrSafe Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As LongPtr, ByVal uFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hWnd As LongPtr, ByRef lpTPMParams As TPMPARAMS) As Long
Private Declare PtrSafe Function MapWindowPoints Lib "user32" (ByVal hWndFrom As LongPtr, ByVal hWndTo As LongPtr, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare PtrSafe Function SendInput Lib "user32" (ByVal nInputs As Long, ByRef pInputs As Any, ByVal cbSize As Long) As Long
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanW" (ByVal cChar As Integer) As Integer
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lData As Long, ByVal wData As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal fFlags As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, ByRef CX As Long, ByRef CY As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef lpMII As MENUITEMINFO) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, ByRef MI As MENUINFO) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal uFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hWnd As Long, ByRef lpTPMParams As TPMPARAMS) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, ByRef pInputs As Any, ByVal cbSize As Long) As Long
#End If
Private Const ICC_BAR_CLASSES As Long = &H20
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
#If VBA7 Then
Private Const HWND_DESKTOP As LongPtr = &H0
#Else
Private Const HWND_DESKTOP As Long = &H0
#End If
Private Const COLOR_MENU As Long = 4
Private Const DST_ICON As Long = &H3
Private Const DST_BITMAP As Long = &H4
Private Const DSS_DISABLED As Long = &H20
Private Const MIM_BACKGROUND As Long = &H2
Private Const MIM_MENUDATA As Long = &H8
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_ID As Long = &H2
Private Const MIIM_STRING As Long = &H40
Private Const MIIM_BITMAP As Long = &H80
Private Const MIIM_FTYPE As Long = &H100
Private Const MFT_SEPARATOR As Long = &H800
Private Const MFS_ENABLED As Long = &H0
Private Const MFS_UNCHECKED As Long = &H0
Private Const MFS_DISABLED As Long = &H3
Private Const MFS_CHECKED As Long = &H8
#If VBA7 Then
Private Const HBMMENU_CALLBACK As LongPtr = (-1)
#Else
Private Const HBMMENU_CALLBACK As Long = (-1)
#End If
Private Const TPM_TOPALIGN As Long = &H0
Private Const TPM_LEFTALIGN As Long = &H0
Private Const TPM_LEFTBUTTON As Long = &H0
Private Const TPM_RIGHTALIGN As Long = &H8
Private Const TPM_VERTICAL As Long = &H40
Private Const TPM_RETURNCMD As Long = &H100
Private Const TPM_LAYOUTRTL As Long = &H8000&
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_EX_TRANSPARENT As Long = &H20
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const SW_HIDE As Long = &H0
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_CANCELMODE As Long = &H1F
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_ENTERMENULOOP As Long = &H211
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_COMMAND As Long = &H111
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINT As Long = &H317, PRF_CLIENT As Long = &H4, PRF_ERASEBKGND As Long = &H8
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_DRAWITEM As Long = &H2B, ODT_MENU As Long = &H1, ODS_DISABLED As Long = &H4
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UPDATEUISTATE As Long = &H128, UIS_SET As Long = 1, UISF_HIDEACCEL As Long = &H2
Private Const WM_USER As Long = &H400
Private Const UM_SETBUTTONCX As Long = (WM_USER + 200)
Private Const TB_ENABLEBUTTON As Long = (WM_USER + 1)
Private Const TB_CHECKBUTTON As Long = (WM_USER + 2)
Private Const TB_PRESSBUTTON As Long = (WM_USER + 3)
Private Const TB_HIDEBUTTON As Long = (WM_USER + 4)
Private Const TB_INDETERMINATE As Long = (WM_USER + 5)
Private Const TB_MARKBUTTON As Long = (WM_USER + 6)
Private Const TB_ISBUTTONENABLED As Long = (WM_USER + 9)
Private Const TB_ISBUTTONCHECKED As Long = (WM_USER + 10)
Private Const TB_ISBUTTONPRESSED As Long = (WM_USER + 11)
Private Const TB_ISBUTTONHIDDEN As Long = (WM_USER + 12)
Private Const TB_ISBUTTONINDETERMINATE As Long = (WM_USER + 13)
Private Const TB_ISBUTTONHIGHLIGHTED As Long = (WM_USER + 14)
Private Const TB_SETSTATE As Long = (WM_USER + 17)
Private Const TB_GETSTATE As Long = (WM_USER + 18)
Private Const TB_ADDBUTTONSA As Long = (WM_USER + 20)
Private Const TB_ADDBUTTONSW As Long = (WM_USER + 68)
Private Const TB_ADDBUTTONS As Long = TB_ADDBUTTONSW
Private Const TB_INSERTBUTTONA As Long = (WM_USER + 21)
Private Const TB_INSERTBUTTONW As Long = (WM_USER + 67)
Private Const TB_INSERTBUTTON As Long = TB_INSERTBUTTONW
Private Const TB_DELETEBUTTON As Long = (WM_USER + 22)
Private Const TB_GETBUTTON As Long = (WM_USER + 23)
Private Const TB_BUTTONCOUNT As Long = (WM_USER + 24)
Private Const TB_COMMANDTOINDEX As Long = (WM_USER + 25)
Private Const TB_SAVERESTOREA As Long = (WM_USER + 26)
Private Const TB_SAVERESTOREW As Long = (WM_USER + 76)
Private Const TB_SAVERESTORE As Long = TB_SAVERESTOREW
Private Const TB_CUSTOMIZE As Long = (WM_USER + 27)
Private Const TB_GETITEMRECT As Long = (WM_USER + 29)
Private Const TB_BUTTONSTRUCTSIZE As Long = (WM_USER + 30)
Private Const TB_SETBUTTONSIZE As Long = (WM_USER + 31)
Private Const TB_SETBITMAPSIZE As Long = (WM_USER + 32)
Private Const TB_AUTOSIZE As Long = (WM_USER + 33)
Private Const TB_GETTOOLTIPS As Long = (WM_USER + 35)
Private Const TB_SETTOOLTIPS As Long = (WM_USER + 36)
Private Const TB_GETROWS As Long = (WM_USER + 40)
Private Const TB_GETBUTTONTEXTA As Long = (WM_USER + 45)
Private Const TB_GETBUTTONTEXTW As Long = (WM_USER + 75)
Private Const TB_GETBUTTONTEXT As Long = TB_GETBUTTONTEXTW
Private Const TB_SETIMAGELIST As Long = (WM_USER + 48)
Private Const TB_GETIMAGELIST As Long = (WM_USER + 49)
Private Const TB_GETRECT As Long = (WM_USER + 51)
Private Const TB_SETHOTIMAGELIST As Long = (WM_USER + 52)
Private Const TB_GETHOTIMAGELIST As Long = (WM_USER + 53)
Private Const TB_SETDISABLEDIMAGELIST As Long = (WM_USER + 54)
Private Const TB_GETDISABLEDIMAGELIST As Long = (WM_USER + 55)
Private Const TB_SETSTYLE As Long = (WM_USER + 56)
Private Const TB_GETSTYLE As Long = (WM_USER + 57)
Private Const TB_GETBUTTONSIZE As Long = (WM_USER + 58)
Private Const TB_SETBUTTONWIDTH As Long = (WM_USER + 59)
Private Const TB_SETMAXTEXTROWS As Long = (WM_USER + 60)
Private Const TB_GETTEXTROWS As Long = (WM_USER + 61)
Private Const TB_GETBUTTONINFOA As Long = (WM_USER + 65)
Private Const TB_GETBUTTONINFOW As Long = (WM_USER + 63)
Private Const TB_GETBUTTONINFO As Long = TB_GETBUTTONINFOW
Private Const TB_SETBUTTONINFOA As Long = (WM_USER + 66)
Private Const TB_SETBUTTONINFOW As Long = (WM_USER + 64)
Private Const TB_SETBUTTONINFO As Long = TB_SETBUTTONINFOW
Private Const TB_HITTEST As Long = (WM_USER + 69)
Private Const TB_SETDRAWTEXTFLAGS = (WM_USER + 70), DT_RTLREADING As Long = &H20000
Private Const TB_GETHOTITEM As Long = (WM_USER + 71)
Private Const TB_SETHOTITEM As Long = (WM_USER + 72)
Private Const TB_SETANCHORHIGHLIGHT As Long = (WM_USER + 73)
Private Const TB_GETANCHORHIGHLIGHT As Long = (WM_USER + 74)
Private Const TB_MAPACCELERATORA As Long = (WM_USER + 78)
Private Const TB_MAPACCELERATORW As Long = (WM_USER + 90)
Private Const TB_MAPACCELERATOR As Long = TB_MAPACCELERATORW
Private Const TB_GETINSERTMARK As Long = (WM_USER + 79)
Private Const TB_SETINSERTMARK As Long = (WM_USER + 80)
Private Const TB_INSERTMARKHITTEST As Long = (WM_USER + 81)
Private Const TB_MOVEBUTTON As Long = (WM_USER + 82)
Private Const TB_GETMAXSIZE As Long = (WM_USER + 83)
Private Const TB_SETEXTENDEDSTYLE As Long = (WM_USER + 84)
Private Const TB_GETEXTENDEDSTYLE As Long = (WM_USER + 85)
Private Const TB_GETPADDING As Long = (WM_USER + 86)
Private Const TB_SETPADDING As Long = (WM_USER + 87)
Private Const TB_SETINSERTMARKCOLOR As Long = (WM_USER + 88)
Private Const TB_GETINSERTMARKCOLOR As Long = (WM_USER + 89)
Private Const TB_GETIDEALSIZE As Long = (WM_USER + 99)
Private Const TB_SETPRESSEDIMAGELIST As Long = (WM_USER + 104)
Private Const TB_GETPRESSEDIMAGELIST As Long = (WM_USER + 105)
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETUNICODEFORMAT As Long = (CCM_FIRST + 5)
Private Const TB_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT
Private Const CCS_TOP As Long = &H1
Private Const CCS_VERT As Long = &H80
Private Const CCS_NORESIZE As Long = &H4
Private Const CCS_NODIVIDER As Long = &H40
Private Const CCS_ADJUSTABLE As Long = &H20
Private Const TBIF_IMAGE As Long = &H1
Private Const TBIF_TEXT As Long = &H2
Private Const TBIF_STATE As Long = &H4
Private Const TBIF_STYLE As Long = &H8
Private Const TBIF_LPARAM As Long = &H10
Private Const TBIF_COMMAND As Long = &H20
Private Const TBIF_SIZE As Long = &H40
Private Const TBIF_BYINDEX As Long = &H80000000
Private Const TBSTATE_CHECKED As Long = &H1
Private Const TBSTATE_PRESSED As Long = &H2
Private Const TBSTATE_ENABLED As Long = &H4
Private Const TBSTATE_HIDDEN As Long = &H8
Private Const TBSTATE_INDETERMINATE As Long = &H10
Private Const TBSTATE_WRAP As Long = &H20
Private Const TBSTATE_ELLIPSES As Long = &H40
Private Const TBSTATE_MARKED As Long = &H80
Private Const TBN_FIRST As Long = (-700)
Private Const TBN_GETBUTTONINFOA As Long = (TBN_FIRST - 0)
Private Const TBN_GETBUTTONINFOW As Long = (TBN_FIRST - 20)
Private Const TBN_GETBUTTONINFO As Long = TBN_GETBUTTONINFOW
Private Const TBN_BEGINDRAG As Long = (TBN_FIRST - 1)
Private Const TBN_ENDDRAG As Long = (TBN_FIRST - 2)
Private Const TBN_BEGINADJUST As Long = (TBN_FIRST - 3)
Private Const TBN_ENDADJUST As Long = (TBN_FIRST - 4)
Private Const TBN_RESET As Long = (TBN_FIRST - 5)
Private Const TBN_QUERYINSERT As Long = (TBN_FIRST - 6)
Private Const TBN_QUERYDELETE As Long = (TBN_FIRST - 7)
Private Const TBN_TOOLBARCHANGE As Long = (TBN_FIRST - 8)
Private Const TBN_CUSTHELP As Long = (TBN_FIRST - 9)
Private Const TBN_DROPDOWN As Long = (TBN_FIRST - 10)
Private Const TBN_GETOBJECT As Long = (TBN_FIRST - 12)
Private Const TBN_HOTITEMCHANGE As Long = (TBN_FIRST - 13)
Private Const TBN_DRAGOUT As Long = (TBN_FIRST - 14)
Private Const TBN_DELETINGBUTTON As Long = (TBN_FIRST - 15)
Private Const TBN_GETDISPINFOA As Long = (TBN_FIRST - 16)
Private Const TBN_GETDISPINFOW As Long = (TBN_FIRST - 17)
Private Const TBN_GETDISPINFO As Long = TBN_GETDISPINFOW
Private Const TBN_GETINFOTIPA As Long = (TBN_FIRST - 18)
Private Const TBN_GETINFOTIPW As Long = (TBN_FIRST - 19)
Private Const TBN_GETINFOTIP As Long = TBN_GETINFOTIPW
Private Const TBN_RESTORE As Long = (TBN_FIRST - 21)
Private Const TBN_SAVE As Long = (TBN_FIRST - 22)
Private Const TBN_INITCUSTOMIZE As Long = (TBN_FIRST - 23)
Private Const TBN_WRAPHOTITEM As Long = (TBN_FIRST - 24)
Private Const TBN_DUPACCELERATOR As Long = (TBN_FIRST - 25)
Private Const TBN_WRAPACCELERATOR As Long = (TBN_FIRST - 26)
Private Const TBN_DRAGOVER As Long = (TBN_FIRST - 27)
Private Const TBN_MAPACCELERATOR As Long = (TBN_FIRST - 28)
Private Const TBNF_IMAGE As Long = &H1
Private Const TBNF_TEXT As Long = &H2
Private Const TBNF_DI_SETITEM As Long = &H10000000
Private Const TBNRF_HIDEHELP As Long = &H1
Private Const TBNRF_ENDCUSTOMIZE As Long = &H2
Private Const TBDDRET_DEFAULT As Long = 0
Private Const TBDDRET_NODEFAULT As Long = 1
Private Const TBDDRET_TREATPRESSED As Long = 2
Private Const TBIMHT_AFTER As Long = &H1
Private Const TBIMHT_BACKGROUND As Long = &H2
Private Const I_IMAGENONE As Long = (-2)
Private Const NM_FIRST As Long = 0
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Private Const NM_TOOLTIPSCREATED As Long = (NM_FIRST - 19)
Private Const BTNS_BUTTON As Long = &H0
Private Const BTNS_SEP As Long = &H1
Private Const BTNS_CHECK As Long = &H2
Private Const BTNS_GROUP As Long = &H4
Private Const BTNS_CHECKGROUP As Long = (BTNS_GROUP Or BTNS_CHECK)
Private Const BTNS_DROPDOWN As Long = &H8
Private Const BTNS_AUTOSIZE As Long = &H10
Private Const BTNS_NOPREFIX As Long = &H20
Private Const BTNS_WHOLEDROPDOWN As Long = &H80
Private Const TBSTYLE_TOOLTIPS As Long = &H100
Private Const TBSTYLE_WRAPABLE As Long = &H200
Private Const TBSTYLE_ALTDRAG As Long = &H400
Private Const TBSTYLE_FLAT As Long = &H800
Private Const TBSTYLE_LIST As Long = &H1000
Private Const TBSTYLE_TRANSPARENT As Long = &H8000&
Private Const TBSTYLE_EX_DRAWDDARROWS As Long = &H1
Private Const TBSTYLE_EX_HIDECLIPPEDBUTTONS As Long = &H10
Private Const TBSTYLE_EX_DOUBLEBUFFER As Long = &H80
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IPerPropertyBrowsingVB
Private Type InitButtonMenuStruct
Key As String
Tag As String
Text As String
Enabled As Boolean
Visible As Boolean
Checked As Boolean
Separator As Boolean
Picture As IPictureDisp
End Type
Private Type InitButtonStruct
Key As String
Tag As String
Caption As String
Style As TbrButtonStyleConstants
Image As Variant
ImageIndex As Long
ToolTipText As String
Description As String
Value As TbrButtonValueConstants
ForeColor As OLE_COLOR
Enabled As Boolean
Visible As Boolean
MixedState As Boolean
HighLighted As Boolean
NoImage As Boolean
NoPrefix As Boolean
AutoSize As Boolean
CustomWidth As Long
ButtonMenusCount As Long
ButtonMenus() As InitButtonMenuStruct
End Type
Private Type ShadowButtonStruct
TBB As TBBUTTON
Caption As String
CX As Long
End Type
Private ToolBarHandle As LongPtr, ToolBarToolTipHandle As LongPtr
Private ToolBarBackColorBrush As LongPtr
Private ToolBarTransparentBrush As LongPtr
Private ToolBarFontHandle As LongPtr
Private ToolBarCustomizeButtonsCount As Long
Private ToolBarCustomizeButtons() As ShadowButtonStruct
Private ToolBarIsClick As Boolean
Private ToolBarMouseOver As Boolean, ToolBarMouseOverIndex As Long
Private ToolBarDesignMode As Boolean
Private ToolBarResizeFrozen As Boolean
Private ToolBarImageSize As Long, ToolBarDefaultImageSize As Long
Private ToolBarDoubleBufferEraseBkgDC As LongPtr
Private ToolBarAlignable As Boolean
Private ToolBarImageListObjectPointer As LongPtr
Private ToolBarDisabledImageListObjectPointer As LongPtr
Private ToolBarHotImageListObjectPointer As LongPtr
Private ToolBarPressedImageListObjectPointer As LongPtr
Private ToolBarPopupMenuHandle As LongPtr, ToolBarPopupMenuButton As TbrButton, ToolBarPopupMenuKeyboard As Boolean
Private DispIDMousePointer As Long
Private DispIDImageList As Long, ImageListArray() As String, ImageListSize As SIZEAPI
Private DispIDDisabledImageList As Long, DisabledImageListArray() As String, DisabledImageListSize As SIZEAPI
Private DispIDHotImageList As Long, HotImageListArray() As String, HotImageListSize As SIZEAPI
Private DispIDPressedImageList As Long, PressedImageListArray() As String, PressedImageListSize As SIZEAPI
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropButtons As TbrButtons
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropImageListName As String, PropImageListInit As Boolean
Private PropDisabledImageListName As String, PropDisabledImageListInit As Boolean
Private PropHotImageListName As String, PropHotImageListInit As Boolean
Private PropPressedImageListName As String, PropPressedImageListInit As Boolean
Private PropBackColor As OLE_COLOR
Private PropStyle As TbrStyleConstants
Private PropTextAlignment As TbrTextAlignConstants
Private PropOrientation As TbrOrientationConstants
Private PropDivider As Boolean
Private PropShowTips As Boolean
Private PropWrappable As Boolean
Private PropAllowCustomize As Boolean
Private PropAltDrag As Boolean
Private PropDoubleBuffer As Boolean
Private PropButtonHeight As Integer
Private PropButtonWidth As Integer
Private PropMinButtonWidth As Integer
Private PropMaxButtonWidth As Integer
Private PropInsertMarkColor As OLE_COLOR
Private PropTransparent As Boolean
Private PropHotTracking As Boolean
Private PropHideClippedButtons As Boolean
Private PropAnchorHot As Boolean
Private PropMaxTextRows As Integer

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
ElseIf DispID = DispIDImageList Then
    DisplayName = PropImageListName
    Handled = True
ElseIf DispID = DispIDDisabledImageList Then
    DisplayName = PropDisabledImageListName
    Handled = True
ElseIf DispID = DispIDHotImageList Then
    DisplayName = PropHotImageListName
    Handled = True
ElseIf DispID = DispIDPressedImageList Then
    DisplayName = PropPressedImageListName
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
    Handled = True
ElseIf DispID = DispIDImageList Or DispID = DispIDDisabledImageList Or DispID = DispIDHotImageList Or DispID = DispIDPressedImageList Then
    On Error GoTo CATCH_EXCEPTION
    Call ComCtlsIPPBSetPredefinedStringsImageList(StringsOut(), CookiesOut(), UserControl.ParentControls, ImageListArray())
    DisabledImageListArray() = ImageListArray()
    HotImageListArray() = ImageListArray()
    PressedImageListArray() = ImageListArray()
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
ElseIf DispID = DispIDDisabledImageList Then
    If Cookie < UBound(DisabledImageListArray()) Then Value = DisabledImageListArray(Cookie)
    Handled = True
ElseIf DispID = DispIDHotImageList Then
    If Cookie < UBound(HotImageListArray()) Then Value = HotImageListArray(Cookie)
    Handled = True
ElseIf DispID = DispIDPressedImageList Then
    If Cookie < UBound(PressedImageListArray()) Then Value = PressedImageListArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_BAR_CLASSES)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim ImageListArray(0) As String
ReDim DisabledImageListArray(0) As String
ReDim HotImageListArray(0) As String
ReDim PressedImageListArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
If DispIDDisabledImageList = 0 Then DispIDDisabledImageList = GetDispID(Me, "DisabledImageList")
If DispIDHotImageList = 0 Then DispIDHotImageList = GetDispID(Me, "HotImageList")
If DispIDPressedImageList = 0 Then DispIDPressedImageList = GetDispID(Me, "PressedImageList")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then ToolBarAlignable = False Else ToolBarAlignable = True
ToolBarDesignMode = Not Ambient.UserMode
On Error GoTo 0
If ToolBarAlignable = True Then Extender.Align = vbAlignTop
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = "(None)"
PropDisabledImageListName = "(None)"
PropHotImageListName = "(None)"
PropPressedImageListName = "(None)"
PropBackColor = vbButtonFace
PropStyle = TbrStyleStandard
PropTextAlignment = TbrTextAlignBottom
PropOrientation = TbrOrientationHorizontal
PropDivider = True
PropShowTips = False
PropWrappable = True
PropAllowCustomize = True
PropAltDrag = False
PropDoubleBuffer = True
PropButtonHeight = (22 * PixelsPerDIP_Y())
PropButtonWidth = (24 * PixelsPerDIP_X())
PropMinButtonWidth = 0
PropMaxButtonWidth = 0
PropInsertMarkColor = vbBlack
PropTransparent = False
PropHotTracking = False
PropHideClippedButtons = False
PropAnchorHot = False
PropMaxTextRows = 1
Call CreateToolBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
If DispIDDisabledImageList = 0 Then DispIDDisabledImageList = GetDispID(Me, "DisabledImageList")
If DispIDHotImageList = 0 Then DispIDHotImageList = GetDispID(Me, "HotImageList")
If DispIDPressedImageList = 0 Then DispIDPressedImageList = GetDispID(Me, "PressedImageList")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then ToolBarAlignable = False Else ToolBarAlignable = True
ToolBarDesignMode = Not Ambient.UserMode
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
PropDisabledImageListName = .ReadProperty("DisabledImageList", "(None)")
PropHotImageListName = .ReadProperty("HotImageList", "(None)")
PropPressedImageListName = .ReadProperty("PressedImageList", "(None)")
PropBackColor = .ReadProperty("BackColor", vbButtonFace)
PropStyle = .ReadProperty("Style", TbrStyleStandard)
PropTextAlignment = .ReadProperty("TextAlignment", TbrTextAlignBottom)
PropOrientation = .ReadProperty("Orientation", TbrOrientationHorizontal)
PropDivider = .ReadProperty("Divider", True)
PropShowTips = .ReadProperty("ShowTips", False)
PropWrappable = .ReadProperty("Wrappable", True)
PropAllowCustomize = .ReadProperty("AllowCustomize", True)
PropAltDrag = .ReadProperty("AltDrag", False)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
PropButtonHeight = (.ReadProperty("ButtonHeight", 22) * PixelsPerDIP_Y())
PropButtonWidth = (.ReadProperty("ButtonWidth", 24) * PixelsPerDIP_X())
PropMinButtonWidth = (.ReadProperty("MinButtonWidth", 0) * PixelsPerDIP_X())
PropMaxButtonWidth = (.ReadProperty("MaxButtonWidth", 0) * PixelsPerDIP_X())
PropInsertMarkColor = .ReadProperty("InsertMarkColor", vbBlack)
PropTransparent = .ReadProperty("Transparent", False)
PropHotTracking = .ReadProperty("HotTracking", False)
PropHideClippedButtons = .ReadProperty("HideClippedButtons", False)
PropAnchorHot = .ReadProperty("AnchorHot", False)
PropMaxTextRows = .ReadProperty("MaxTextRows", 1)
End With
With New PropertyBag
On Error Resume Next
.Contents = PropBag.ReadProperty("InitButtons", 0)
On Error GoTo 0
Dim InitButtonsCount As Long, i As Long, ii As Long
Dim InitButtons() As InitButtonStruct
InitButtonsCount = .ReadProperty("InitButtonsCount", 0)
If InitButtonsCount > 0 Then
    ReDim InitButtons(1 To InitButtonsCount) As InitButtonStruct
    Dim VarValue As Variant, PropBagButtonMenus As PropertyBag
    For i = 1 To InitButtonsCount
        InitButtons(i).Key = VarToStr(.ReadProperty("InitButtonsKey" & CStr(i), vbNullString))
        InitButtons(i).Tag = VarToStr(.ReadProperty("InitButtonsTag" & CStr(i), vbNullString))
        InitButtons(i).Caption = VarToStr(.ReadProperty("InitButtonsCaption" & CStr(i), vbNullString))
        InitButtons(i).Style = .ReadProperty("InitButtonsStyle" & CStr(i), TbrButtonStyleDefault)
        VarValue = .ReadProperty("InitButtonsImage" & CStr(i), 0)
        If VarType(VarValue) = vbArray + vbByte Then
            InitButtons(i).Image = VarToStr(VarValue)
            InitButtons(i).ImageIndex = .ReadProperty("InitButtonsImageIndex" & CStr(i), 0)
        Else
            InitButtons(i).Image = VarValue
            InitButtons(i).ImageIndex = CLng(VarValue)
        End If
        InitButtons(i).ToolTipText = VarToStr(.ReadProperty("InitButtonsToolTipText" & CStr(i), vbNullString))
        InitButtons(i).Description = VarToStr(.ReadProperty("InitButtonsDescription" & CStr(i), vbNullString))
        InitButtons(i).Value = .ReadProperty("InitButtonsValue" & CStr(i), TbrButtonValueUnpressed)
        InitButtons(i).ForeColor = .ReadProperty("InitButtonsForeColor" & CStr(i), vbButtonText)
        InitButtons(i).Enabled = .ReadProperty("InitButtonsEnabled" & CStr(i), True)
        InitButtons(i).Visible = .ReadProperty("InitButtonsVisible" & CStr(i), True)
        InitButtons(i).MixedState = .ReadProperty("InitButtonsMixedState" & CStr(i), False)
        InitButtons(i).HighLighted = .ReadProperty("InitButtonsHighLighted" & CStr(i), False)
        InitButtons(i).NoImage = .ReadProperty("InitButtonsNoImage" & CStr(i), False)
        InitButtons(i).NoPrefix = .ReadProperty("InitButtonsNoPrefix" & CStr(i), False)
        InitButtons(i).AutoSize = .ReadProperty("InitButtonsAutoSize" & CStr(i), False)
        InitButtons(i).CustomWidth = (.ReadProperty("InitButtonsCustomWidth" & CStr(i), 0) * PixelsPerDIP_X())
        Set PropBagButtonMenus = New PropertyBag
        PropBagButtonMenus.Contents = .ReadProperty("InitButtonsButtonMenus" & CStr(i), 0)
        InitButtons(i).ButtonMenusCount = PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusCount", 0)
        If InitButtons(i).ButtonMenusCount > 0 Then
            ReDim InitButtons(i).ButtonMenus(1 To InitButtons(i).ButtonMenusCount) As InitButtonMenuStruct
            For ii = 1 To InitButtons(i).ButtonMenusCount
                InitButtons(i).ButtonMenus(ii).Key = VarToStr(PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusKey" & CStr(ii), vbNullString))
                InitButtons(i).ButtonMenus(ii).Tag = VarToStr(PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusTag" & CStr(ii), vbNullString))
                InitButtons(i).ButtonMenus(ii).Text = VarToStr(PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusText" & CStr(ii), vbNullString))
                InitButtons(i).ButtonMenus(ii).Enabled = PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusEnabled" & CStr(ii), True)
                InitButtons(i).ButtonMenus(ii).Visible = PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusVisible" & CStr(ii), True)
                InitButtons(i).ButtonMenus(ii).Checked = PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusChecked" & CStr(ii), False)
                InitButtons(i).ButtonMenus(ii).Separator = PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusSeparator" & CStr(ii), False)
                Set InitButtons(i).ButtonMenus(ii).Picture = PropBagButtonMenus.ReadProperty("InitButtonsButtonMenusPicture" & CStr(ii), Nothing)
            Next ii
        End If
    Next i
End If
End With
Call CreateToolBar
If InitButtonsCount > 0 And ToolBarHandle <> NULL_PTR Then
    Dim ImageListInit As Boolean
    ImageListInit = PropImageListInit
    PropImageListInit = True
    For i = 1 To InitButtonsCount
        With Me.Buttons.Add(i, InitButtons(i).Key, InitButtons(i).Caption, InitButtons(i).Style, InitButtons(i).ImageIndex)
        .FInit Me, InitButtons(i).Key, InitButtons(i).Caption, InitButtons(i).Image, InitButtons(i).ImageIndex
        .Tag = InitButtons(i).Tag
        .ToolTipText = InitButtons(i).ToolTipText
        .Description = InitButtons(i).Description
        If InitButtons(i).Value = TbrButtonValuePressed Then .Value = TbrButtonValuePressed
        .ForeColor = InitButtons(i).ForeColor
        If InitButtons(i).Enabled = False Then .Enabled = False
        If InitButtons(i).Visible = False Then .Visible = False
        If InitButtons(i).MixedState = True Then .MixedState = True
        If InitButtons(i).HighLighted = True Then .HighLighted = True
        If InitButtons(i).NoImage = True Then .NoImage = True
        If InitButtons(i).NoPrefix = True Then .NoPrefix = True
        If InitButtons(i).AutoSize = True Then .AutoSize = True
        If InitButtons(i).CustomWidth > 0 Then .CustomWidth = UserControl.ScaleX(InitButtons(i).CustomWidth, vbPixels, vbContainerSize)
        If InitButtons(i).ButtonMenusCount > 0 Then
            For ii = 1 To InitButtons(i).ButtonMenusCount
                With .ButtonMenus.Add(ii, InitButtons(i).ButtonMenus(ii).Key, InitButtons(i).ButtonMenus(ii).Text)
                .Tag = InitButtons(i).ButtonMenus(ii).Tag
                If InitButtons(i).ButtonMenus(ii).Enabled = False Then .Enabled = False
                If InitButtons(i).ButtonMenus(ii).Visible = False Then .Visible = False
                If InitButtons(i).ButtonMenus(ii).Checked = True Then .Checked = True
                If InitButtons(i).ButtonMenus(ii).Separator = True Then .Separator = True
                If Not InitButtons(i).ButtonMenus(ii).Picture Is Nothing Then Set .Picture = InitButtons(i).ButtonMenus(ii).Picture
                End With
            Next ii
        End If
        End With
    Next i
    PropImageListInit = ImageListInit
End If
If Not PropImageListName = "(None)" Or Not PropDisabledImageListName = "(None)" Or Not PropHotImageListName = "(None)" Or Not PropPressedImageListName = "(None)" Then TimerImageList.Enabled = True
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
.WriteProperty "DisabledImageList", PropDisabledImageListName, "(None)"
.WriteProperty "HotImageList", PropHotImageListName, "(None)"
.WriteProperty "PressedImageList", PropPressedImageListName, "(None)"
.WriteProperty "BackColor", PropBackColor, vbButtonFace
.WriteProperty "Style", PropStyle, TbrStyleStandard
.WriteProperty "TextAlignment", PropTextAlignment, TbrTextAlignBottom
.WriteProperty "Orientation", PropOrientation, TbrOrientationHorizontal
.WriteProperty "Divider", PropDivider, True
.WriteProperty "ShowTips", PropShowTips, False
.WriteProperty "Wrappable", PropWrappable, True
.WriteProperty "AllowCustomize", PropAllowCustomize, True
.WriteProperty "AltDrag", PropAltDrag, False
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
.WriteProperty "ButtonHeight", (PropButtonHeight / PixelsPerDIP_Y()), 22
.WriteProperty "ButtonWidth", (PropButtonWidth / PixelsPerDIP_X()), 24
.WriteProperty "MinButtonWidth", (PropMinButtonWidth / PixelsPerDIP_X()), 0
.WriteProperty "MaxButtonWidth", (PropMaxButtonWidth / PixelsPerDIP_X()), 0
.WriteProperty "InsertMarkColor", PropInsertMarkColor, vbBlack
.WriteProperty "Transparent", PropTransparent, False
.WriteProperty "HotTracking", PropHotTracking, False
.WriteProperty "HideClippedButtons", PropHideClippedButtons, False
.WriteProperty "AnchorHot", PropAnchorHot, False
.WriteProperty "MaxTextRows", PropMaxTextRows, 1
End With
Dim Count(0 To 1) As Long
Count(0) = Me.Buttons.Count
With New PropertyBag
.WriteProperty "InitButtonsCount", Count(0), 0
If Count(0) > 0 Then
    Dim i As Long, ii As Long, VarValue As Variant, PropBagButtonMenus As PropertyBag
    For i = 1 To Count(0)
        .WriteProperty "InitButtonsKey" & CStr(i), StrToVar(Me.Buttons(i).Key), vbNullString
        .WriteProperty "InitButtonsTag" & CStr(i), StrToVar(Me.Buttons(i).Tag), vbNullString
        .WriteProperty "InitButtonsCaption" & CStr(i), StrToVar(Me.Buttons(i).Caption), vbNullString
        .WriteProperty "InitButtonsStyle" & CStr(i), Me.Buttons(i).Style, TbrButtonStyleDefault
        VarValue = Me.Buttons(i).Image
        If VarType(VarValue) = vbString Then
            .WriteProperty "InitButtonsImage" & CStr(i), StrToVar(VarValue), 0
            .WriteProperty "InitButtonsImageIndex" & CStr(i), Me.Buttons(i).ImageIndex, 0
        Else
            .WriteProperty "InitButtonsImage" & CStr(i), VarValue, 0
        End If
        .WriteProperty "InitButtonsToolTipText" & CStr(i), StrToVar(Me.Buttons(i).ToolTipText), vbNullString
        .WriteProperty "InitButtonsDescription" & CStr(i), StrToVar(Me.Buttons(i).Description), vbNullString
        .WriteProperty "InitButtonsValue" & CStr(i), Me.Buttons(i).Value, TbrButtonValueUnpressed
        .WriteProperty "InitButtonsForeColor" & CStr(i), Me.Buttons(i).ForeColor, vbButtonText
        .WriteProperty "InitButtonsEnabled" & CStr(i), Me.Buttons(i).Enabled, True
        .WriteProperty "InitButtonsVisible" & CStr(i), Me.Buttons(i).Visible, True
        .WriteProperty "InitButtonsMixedState" & CStr(i), Me.Buttons(i).MixedState, False
        .WriteProperty "InitButtonsHighLighted" & CStr(i), Me.Buttons(i).HighLighted, False
        .WriteProperty "InitButtonsNoImage" & CStr(i), Me.Buttons(i).NoImage, False
        .WriteProperty "InitButtonsNoPrefix" & CStr(i), Me.Buttons(i).NoPrefix, False
        .WriteProperty "InitButtonsAutoSize" & CStr(i), Me.Buttons(i).AutoSize, False
        .WriteProperty "InitButtonsCustomWidth" & CStr(i), (CLng(UserControl.ScaleX(Me.Buttons(i).CustomWidth, vbContainerSize, vbPixels)) / PixelsPerDIP_X()), 0
        Set PropBagButtonMenus = New PropertyBag
        Count(1) = Me.Buttons(i).ButtonMenus.Count
        PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusCount", Count(1), 0
        If Count(1) > 0 Then
            For ii = 1 To Count(1)
                PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusKey" & CStr(ii), StrToVar(Me.Buttons(i).ButtonMenus(ii).Key), vbNullString
                PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusTag" & CStr(ii), StrToVar(Me.Buttons(i).ButtonMenus(ii).Tag), vbNullString
                PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusText" & CStr(ii), StrToVar(Me.Buttons(i).ButtonMenus(ii).Text), vbNullString
                PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusEnabled" & CStr(ii), Me.Buttons(i).ButtonMenus(ii).Enabled, True
                PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusVisible" & CStr(ii), Me.Buttons(i).ButtonMenus(ii).Visible, True
                PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusChecked" & CStr(ii), Me.Buttons(i).ButtonMenus(ii).Checked, False
                PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusSeparator" & CStr(ii), Me.Buttons(i).ButtonMenus(ii).Separator, False
                PropBagButtonMenus.WriteProperty "InitButtonsButtonMenusPicture" & CStr(ii), Me.Buttons(i).ButtonMenus(ii).Picture, Nothing
            Next ii
        End If
        .WriteProperty "InitButtonsButtonMenus" & CStr(i), PropBagButtonMenus.Contents, 0
    Next i
End If
PropBag.WriteProperty "InitButtons", .Contents, 0
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
If InProc = True Or ToolBarResizeFrozen = True Then Exit Sub
InProc = True
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If ToolBarHandle = NULL_PTR Then InProc = False: Exit Sub
SendMessage ToolBarHandle, TB_AUTOSIZE, 0, ByVal 0&
Dim dwStyle As Long, Count As Long, Size As SIZEAPI, Rows As Long
dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
Count = CLng(SendMessage(ToolBarHandle, TB_BUTTONCOUNT, 0, ByVal 0&))
If Count > 0 Then
    Dim i As Long, RC As RECT
    If PropWrappable = False Then
        If ToolBarAlignable = True Then
            Select Case Extender.Align
                Case vbAlignNone
                    If PropOrientation = TbrOrientationVertical Then Count = 1
                Case vbAlignLeft, vbAlignRight
                    Count = 1
            End Select
        Else
            If PropOrientation = TbrOrientationVertical Then Count = 1
        End If
    End If
    For i = 0 To Count - 1
        If SendMessage(ToolBarHandle, TB_GETITEMRECT, i, ByVal VarPtr(RC)) <> 0 Then
            If RC.Right > Size.CX Then Size.CX = RC.Right
            If RC.Bottom > Size.CY Then Size.CY = RC.Bottom
        End If
    Next i
    If PropWrappable = True Then Rows = CLng(SendMessage(ToolBarHandle, TB_GETROWS, 0, ByVal 0&))
Else
    Size.CX = PropButtonWidth
    Size.CY = PropButtonHeight
End If
If Not (dwStyle And CCS_NODIVIDER) = CCS_NODIVIDER Then
    ' The divider line is a two-pixel highlight.
    Size.CY = Size.CY + 2
End If
With UserControl
Dim Align As Integer
If ToolBarAlignable = True Then Align = .Extender.Align Else Align = vbAlignNone
Select Case Align
    Case vbAlignNone
        If PropOrientation = TbrOrientationHorizontal Then
            On Error Resume Next
            .Extender.Height = .ScaleY(Size.CY, vbPixels, vbContainerSize)
            On Error GoTo 0
            If (dwStyle And CCS_VERT) = CCS_VERT Then
                SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle And Not CCS_VERT
                InProc = False
                Call ReCreateToolBar
                Exit Sub
            End If
        ElseIf PropOrientation = TbrOrientationVertical Then
            On Error Resume Next
            .Extender.Width = .ScaleX(Size.CX, vbPixels, vbContainerSize)
            On Error GoTo 0
            If Not (dwStyle And CCS_VERT) = CCS_VERT Then
                SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle Or CCS_VERT
                InProc = False
                Call ReCreateToolBar
                Exit Sub
            End If
        End If
    Case vbAlignTop, vbAlignBottom
        .Extender.Height = .ScaleY(Size.CY, vbPixels, vbContainerSize)
        If (dwStyle And CCS_VERT) = CCS_VERT Then
            PropOrientation = TbrOrientationHorizontal
            SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle And Not CCS_VERT
            InProc = False
            Call ReCreateToolBar
            Exit Sub
        End If
    Case vbAlignLeft, vbAlignRight
        .Extender.Width = .ScaleX(Size.CX, vbPixels, vbContainerSize)
        If Not (dwStyle And CCS_VERT) = CCS_VERT Then
            PropOrientation = TbrOrientationVertical
            SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle Or CCS_VERT
            InProc = False
            Call ReCreateToolBar
            Exit Sub
        End If
End Select
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If PropTransparent = True Then
    MoveWindow ToolBarHandle, 0, 0, .ScaleWidth, .ScaleHeight, 0
    If ToolBarTransparentBrush <> NULL_PTR Then
        DeleteObject ToolBarTransparentBrush
        ToolBarTransparentBrush = NULL_PTR
    End If
    RedrawWindow ToolBarHandle, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
Else
    MoveWindow ToolBarHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End If
End With
InProc = False
If Count > 0 And PropWrappable = True Then
    If Rows <> SendMessage(ToolBarHandle, TB_GETROWS, 0, ByVal 0&) Then Call UserControl_Resize
End If
End Sub

Private Sub UserControl_Show()
If ToolBarDesignMode = True Then Call UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyToolBar
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropImageListInit = False Then
    If Not PropImageListName = "(None)" Then Me.ImageList = PropImageListName
    PropImageListInit = True
End If
If PropDisabledImageListInit = False Then
    If Not PropDisabledImageListName = "(None)" Then Me.DisabledImageList = PropDisabledImageListName
    PropDisabledImageListInit = True
End If
If PropHotImageListInit = False Then
    If Not PropHotImageListName = "(None)" Then Me.HotImageList = PropHotImageListName
    PropHotImageListInit = True
End If
If PropPressedImageListInit = False Then
    If Not PropPressedImageListName = "(None)" Then Me.PressedImageList = PropPressedImageListName
    PropPressedImageListInit = True
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

#If VBA7 Then
Public Property Get hWnd() As LongPtr
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
#Else
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a control."
Attribute hWnd.VB_UserMemId = -515
#End If
hWnd = ToolBarHandle
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
OldFontHandle = ToolBarFontHandle
ToolBarFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ToolBarHandle <> NULL_PTR Then SendMessage ToolBarHandle, WM_SETFONT, ToolBarFontHandle, ByVal 1&
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
Call ReCreateButtons
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As LongPtr
OldFontHandle = ToolBarFontHandle
ToolBarFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ToolBarHandle <> NULL_PTR Then SendMessage ToolBarHandle, WM_SETFONT, ToolBarFontHandle, ByVal 1&
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
Call ReCreateButtons
Call UserControl_Resize
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If ToolBarHandle <> NULL_PTR And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles ToolBarHandle
    Else
        RemoveVisualStyles ToolBarHandle
    End If
    Call SetVisualStylesToolTip
    ' The font need to be set again if the comctl32.dll version is 6.0 or higher. (Bug?)
    If ToolBarFontHandle <> 0 Then SendMessage ToolBarHandle, WM_SETFONT, ToolBarFontHandle, ByVal 1&
    Call UserControl_Resize
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
If ToolBarHandle <> NULL_PTR Then
    EnableWindow ToolBarHandle, IIf(Value = True, 1, 0)
    Me.Refresh
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
If ToolBarDesignMode = False Then Call RefreshMousePointer
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
        If ToolBarDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ToolBarDesignMode = False Then Call RefreshMousePointer
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
If ToolBarDesignMode = False Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
    dwMask = 0
End If
If ToolBarHandle <> NULL_PTR Then Call ReCreateToolBar
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
If ToolBarDesignMode = False Then
    If PropImageListInit = False And ToolBarImageListObjectPointer = NULL_PTR Then
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
If ToolBarHandle <> NULL_PTR Then
    Dim Success As Boolean, Handle As LongPtr, OldSize As SIZEAPI
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
        End If
        If Success = True Then
            LSet OldSize = ImageListSize
            ImageList_GetIconSize Handle, ImageListSize.CX, ImageListSize.CY
            If ImageListSizesAreEqual() = True Then
                SendMessage ToolBarHandle, TB_SETIMAGELIST, 0, ByVal Handle
                ToolBarImageListObjectPointer = ObjPtr(Value)
                PropImageListName = ProperControlName(Value)
            Else
                LSet ImageListSize = OldSize
                If ToolBarDesignMode = True Then
                    MsgBox "ImageList Image sizes must be the same", vbCritical + vbOKOnly
                    Exit Property
                Else
                    Err.Raise Number:=380, Description:="ImageList Image sizes must be the same"
                End If
            End If
        End If
    ElseIf VarType(Value) = vbString Then
        Dim ControlEnum As Object, CompareName As String
        For Each ControlEnum In UserControl.ParentControls
            If TypeName(ControlEnum) = "ImageList" Then
                CompareName = ProperControlName(ControlEnum)
                If CompareName = Value And Not CompareName = vbNullString Then
                    Err.Clear
                    Handle = ControlEnum.hImageList
                    Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
                    If Success = True Then
                        LSet OldSize = ImageListSize
                        ImageList_GetIconSize Handle, ImageListSize.CX, ImageListSize.CY
                        If ImageListSizesAreEqual() = True Then
                            SendMessage ToolBarHandle, TB_SETIMAGELIST, 0, ByVal Handle
                            If ToolBarDesignMode = False Then ToolBarImageListObjectPointer = ObjPtr(ControlEnum)
                            PropImageListName = Value
                            Exit For
                        Else
                            LSet ImageListSize = OldSize
                            If ToolBarDesignMode = True Then
                                MsgBox "ImageList Image sizes must be the same", vbCritical + vbOKOnly
                                Exit Property
                            Else
                                Err.Raise Number:=380, Description:="ImageList Image sizes must be the same"
                            End If
                        End If
                    ElseIf ToolBarDesignMode = True Then
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
        If SendMessage(ToolBarHandle, TB_GETIMAGELIST, 0, ByVal 0&) <> 0 Then SendMessage ToolBarHandle, TB_SETIMAGELIST, 0, ByVal 0&
        ToolBarImageListObjectPointer = NULL_PTR
        PropImageListName = "(None)"
        ImageListSize.CX = 0: ImageListSize.CY = 0
    ElseIf Handle = NULL_PTR Then
        If SendMessage(ToolBarHandle, TB_GETIMAGELIST, 0, ByVal 0&) <> 0 Then SendMessage ToolBarHandle, TB_SETIMAGELIST, 0, ByVal 0&
        ImageListSize.CX = 0: ImageListSize.CY = 0
    End If
    If SendMessage(ToolBarHandle, TB_GETIMAGELIST, 0, ByVal 0&) = 0 Then
        SendMessage ToolBarHandle, TB_SETBITMAPSIZE, 0, ByVal ToolBarDefaultImageSize
        ToolBarImageSize = ToolBarDefaultImageSize
    Else
        ToolBarImageSize = MakeDWord(ImageListSize.CX, ImageListSize.CY)
    End If
    Call ReCreateButtons
    Call UserControl_Resize
End If
UserControl.PropertyChanged "ImageList"
End Property

Public Property Get DisabledImageList() As Variant
Attribute DisabledImageList.VB_Description = "Returns/sets the image list control to be used for disabled buttons."
If ToolBarDesignMode = False Then
    If PropDisabledImageListInit = False And ToolBarDisabledImageListObjectPointer = NULL_PTR Then
        If Not PropDisabledImageListName = "(None)" Then Me.DisabledImageList = PropDisabledImageListName
        PropDisabledImageListInit = True
    End If
    Set DisabledImageList = PropDisabledImageListControl
Else
    DisabledImageList = PropDisabledImageListName
End If
End Property

Public Property Set DisabledImageList(ByVal Value As Variant)
Me.DisabledImageList = Value
End Property

Public Property Let DisabledImageList(ByVal Value As Variant)
If ToolBarHandle <> NULL_PTR Then
    Dim Success As Boolean, Handle As LongPtr, OldSize As SIZEAPI
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
        End If
        If Success = True Then
            LSet OldSize = DisabledImageListSize
            ImageList_GetIconSize Handle, DisabledImageListSize.CX, DisabledImageListSize.CY
            If ImageListSizesAreEqual() = True Then
                SendMessage ToolBarHandle, TB_SETDISABLEDIMAGELIST, 0, ByVal Handle
                ToolBarDisabledImageListObjectPointer = ObjPtr(Value)
                PropDisabledImageListName = ProperControlName(Value)
            Else
                LSet DisabledImageListSize = OldSize
                If ToolBarDesignMode = True Then
                    MsgBox "ImageList Image sizes must be the same", vbCritical + vbOKOnly
                    Exit Property
                Else
                    Err.Raise Number:=380, Description:="ImageList Image sizes must be the same"
                End If
            End If
        End If
    ElseIf VarType(Value) = vbString Then
        Dim ControlEnum As Object, CompareName As String
        For Each ControlEnum In UserControl.ParentControls
            If TypeName(ControlEnum) = "ImageList" Then
                CompareName = ProperControlName(ControlEnum)
                If CompareName = Value And Not CompareName = vbNullString Then
                    Err.Clear
                    Handle = ControlEnum.hImageList
                    Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
                    If Success = True Then
                        LSet OldSize = DisabledImageListSize
                        ImageList_GetIconSize Handle, DisabledImageListSize.CX, DisabledImageListSize.CY
                        If ImageListSizesAreEqual() = True Then
                            SendMessage ToolBarHandle, TB_SETDISABLEDIMAGELIST, 0, ByVal Handle
                            If ToolBarDesignMode = False Then ToolBarDisabledImageListObjectPointer = ObjPtr(ControlEnum)
                            PropDisabledImageListName = Value
                            Exit For
                        Else
                            LSet DisabledImageListSize = OldSize
                            If ToolBarDesignMode = True Then
                                MsgBox "ImageList Image sizes must be the same", vbCritical + vbOKOnly
                                Exit Property
                            Else
                                Err.Raise Number:=380, Description:="ImageList Image sizes must be the same"
                            End If
                        End If
                    ElseIf ToolBarDesignMode = True Then
                        PropDisabledImageListName = Value
                        Success = True
                        Exit For
                    End If
                End If
            End If
        Next ControlEnum
    End If
    On Error GoTo 0
    If Success = False Then
        If SendMessage(ToolBarHandle, TB_GETDISABLEDIMAGELIST, 0, ByVal 0&) <> 0 Then SendMessage ToolBarHandle, TB_SETDISABLEDIMAGELIST, 0, ByVal 0&
        ToolBarDisabledImageListObjectPointer = NULL_PTR
        PropDisabledImageListName = "(None)"
        DisabledImageListSize.CX = 0: DisabledImageListSize.CY = 0
    ElseIf Handle = NULL_PTR Then
        If SendMessage(ToolBarHandle, TB_GETDISABLEDIMAGELIST, 0, ByVal 0&) <> 0 Then SendMessage ToolBarHandle, TB_SETDISABLEDIMAGELIST, 0, ByVal 0&
        DisabledImageListSize.CX = 0: DisabledImageListSize.CY = 0
    End If
    Me.Refresh
End If
UserControl.PropertyChanged "DisabledImageList"
End Property

Public Property Get HotImageList() As Variant
Attribute HotImageList.VB_Description = "Returns/sets the image list control to be used for hot buttons."
If ToolBarDesignMode = False Then
    If PropHotImageListInit = False And ToolBarHotImageListObjectPointer = NULL_PTR Then
        If Not PropHotImageListName = "(None)" Then Me.HotImageList = PropHotImageListName
        PropHotImageListInit = True
    End If
    Set HotImageList = PropHotImageListControl
Else
    HotImageList = PropHotImageListName
End If
End Property

Public Property Set HotImageList(ByVal Value As Variant)
Me.HotImageList = Value
End Property

Public Property Let HotImageList(ByVal Value As Variant)
If ToolBarHandle <> NULL_PTR Then
    Dim Success As Boolean, Handle As LongPtr, OldSize As SIZEAPI
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
        End If
        If Success = True Then
            LSet OldSize = HotImageListSize
            ImageList_GetIconSize Handle, HotImageListSize.CX, HotImageListSize.CY
            If ImageListSizesAreEqual() = True Then
                SendMessage ToolBarHandle, TB_SETHOTIMAGELIST, 0, ByVal Handle
                ToolBarHotImageListObjectPointer = ObjPtr(Value)
                PropHotImageListName = ProperControlName(Value)
            Else
                LSet HotImageListSize = OldSize
                If ToolBarDesignMode = True Then
                    MsgBox "ImageList Image sizes must be the same", vbCritical + vbOKOnly
                    Exit Property
                Else
                    Err.Raise Number:=380, Description:="ImageList Image sizes must be the same"
                End If
            End If
        End If
    ElseIf VarType(Value) = vbString Then
        Dim ControlEnum As Object, CompareName As String
        For Each ControlEnum In UserControl.ParentControls
            If TypeName(ControlEnum) = "ImageList" Then
                CompareName = ProperControlName(ControlEnum)
                If CompareName = Value And Not CompareName = vbNullString Then
                    Err.Clear
                    Handle = ControlEnum.hImageList
                    Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
                    If Success = True Then
                        LSet OldSize = HotImageListSize
                        ImageList_GetIconSize Handle, HotImageListSize.CX, HotImageListSize.CY
                        If ImageListSizesAreEqual() = True Then
                            SendMessage ToolBarHandle, TB_SETHOTIMAGELIST, 0, ByVal Handle
                            If ToolBarDesignMode = False Then ToolBarHotImageListObjectPointer = ObjPtr(ControlEnum)
                            PropHotImageListName = Value
                            Exit For
                        Else
                            LSet HotImageListSize = OldSize
                            If ToolBarDesignMode = True Then
                                MsgBox "ImageList Image sizes must be the same", vbCritical + vbOKOnly
                                Exit Property
                            Else
                                Err.Raise Number:=380, Description:="ImageList Image sizes must be the same"
                            End If
                        End If
                    ElseIf ToolBarDesignMode = True Then
                        PropHotImageListName = Value
                        Success = True
                        Exit For
                    End If
                End If
            End If
        Next ControlEnum
    End If
    On Error GoTo 0
    If Success = False Then
        If SendMessage(ToolBarHandle, TB_GETHOTIMAGELIST, 0, ByVal 0&) <> 0 Then SendMessage ToolBarHandle, TB_SETHOTIMAGELIST, 0, ByVal 0&
        PropHotImageListName = "(None)"
        ToolBarHotImageListObjectPointer = NULL_PTR
        HotImageListSize.CX = 0: HotImageListSize.CY = 0
    ElseIf Handle = NULL_PTR Then
        If SendMessage(ToolBarHandle, TB_GETHOTIMAGELIST, 0, ByVal 0&) <> 0 Then SendMessage ToolBarHandle, TB_SETHOTIMAGELIST, 0, ByVal 0&
        HotImageListSize.CX = 0: HotImageListSize.CY = 0
    End If
    Me.Refresh
End If
UserControl.PropertyChanged "HotImageList"
End Property

Public Property Get PressedImageList() As Variant
Attribute PressedImageList.VB_Description = "Returns/sets the image list control to be used for pressed buttons. Requires comctl32.dll version 6.1 or higher."
If ToolBarDesignMode = False Then
    If PropPressedImageListInit = False And ToolBarPressedImageListObjectPointer = NULL_PTR Then
        If Not PropPressedImageListName = "(None)" Then Me.PressedImageList = PropPressedImageListName
        PropPressedImageListInit = True
    End If
    Set PressedImageList = PropPressedImageListControl
Else
    PressedImageList = PropPressedImageListName
End If
End Property

Public Property Set PressedImageList(ByVal Value As Variant)
Me.PressedImageList = Value
End Property

Public Property Let PressedImageList(ByVal Value As Variant)
If ToolBarHandle <> NULL_PTR Then
    Dim Success As Boolean, Handle As LongPtr, OldSize As SIZEAPI
    On Error Resume Next
    If IsObject(Value) Then
        If TypeName(Value) = "ImageList" Then
            Handle = Value.hImageList
            Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
        End If
        If Success = True Then
            LSet OldSize = PressedImageListSize
            ImageList_GetIconSize Handle, PressedImageListSize.CX, PressedImageListSize.CY
            If ImageListSizesAreEqual() = True Then
                If ComCtlsSupportLevel() >= 2 Then SendMessage ToolBarHandle, TB_SETPRESSEDIMAGELIST, 0, ByVal Handle
                ToolBarPressedImageListObjectPointer = ObjPtr(Value)
                PropPressedImageListName = ProperControlName(Value)
            Else
                LSet PressedImageListSize = OldSize
                If ToolBarDesignMode = True Then
                    MsgBox "ImageList Image sizes must be the same", vbCritical + vbOKOnly
                    Exit Property
                Else
                    Err.Raise Number:=380, Description:="ImageList Image sizes must be the same"
                End If
            End If
        End If
    ElseIf VarType(Value) = vbString Then
        Dim ControlEnum As Object, CompareName As String
        For Each ControlEnum In UserControl.ParentControls
            If TypeName(ControlEnum) = "ImageList" Then
                CompareName = ProperControlName(ControlEnum)
                If CompareName = Value And Not CompareName = vbNullString Then
                    Err.Clear
                    Handle = ControlEnum.hImageList
                    Success = CBool(Err.Number = 0 And Handle <> NULL_PTR)
                    If Success = True Then
                        LSet OldSize = PressedImageListSize
                        ImageList_GetIconSize Handle, PressedImageListSize.CX, PressedImageListSize.CY
                        If ImageListSizesAreEqual() = True Then
                            If ComCtlsSupportLevel() >= 2 Then SendMessage ToolBarHandle, TB_SETPRESSEDIMAGELIST, 0, ByVal Handle
                            PropPressedImageListName = Value
                            If ToolBarDesignMode = False Then ToolBarPressedImageListObjectPointer = ObjPtr(ControlEnum)
                            Exit For
                        Else
                            LSet PressedImageListSize = OldSize
                            If ToolBarDesignMode = True Then
                                MsgBox "ImageList Image sizes must be the same", vbCritical + vbOKOnly
                                Exit Property
                            Else
                                Err.Raise Number:=380, Description:="ImageList Image sizes must be the same"
                            End If
                        End If
                    ElseIf ToolBarDesignMode = True Then
                        PropPressedImageListName = Value
                        Success = True
                        Exit For
                    End If
                End If
            End If
        Next ControlEnum
    End If
    On Error GoTo 0
    If Success = False Then
        If ComCtlsSupportLevel() >= 2 Then If SendMessage(ToolBarHandle, TB_GETPRESSEDIMAGELIST, 0, ByVal 0&) <> 0 Then SendMessage ToolBarHandle, TB_SETPRESSEDIMAGELIST, 0, ByVal 0&
        ToolBarPressedImageListObjectPointer = NULL_PTR
        PropPressedImageListName = "(None)"
        PressedImageListSize.CX = 0: PressedImageListSize.CY = 0
    ElseIf Handle = NULL_PTR Then
        If ComCtlsSupportLevel() >= 2 Then If SendMessage(ToolBarHandle, TB_GETPRESSEDIMAGELIST, 0, ByVal 0&) <> 0 Then SendMessage ToolBarHandle, TB_SETPRESSEDIMAGELIST, 0, ByVal 0&
        PressedImageListSize.CX = 0: PressedImageListSize.CY = 0
    End If
    Me.Refresh
End If
UserControl.PropertyChanged "PressedImageList"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If ToolBarDesignMode = False Then
    If ToolBarHandle <> NULL_PTR Then
        If ToolBarBackColorBrush <> NULL_PTR Then DeleteObject ToolBarBackColorBrush
        ToolBarBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
    End If
End If
UserControl.BackColor = PropBackColor
Me.Refresh
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get Style() As TbrStyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines how the tool bar is drawn."
Style = PropStyle
End Property

Public Property Let Style(ByVal Value As TbrStyleConstants)
Select Case Value
    Case TbrStyleStandard, TbrStyleFlat
        PropStyle = Value
    Case Else
        Err.Raise 380
End Select
If ToolBarHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
    Select Case PropStyle
        Case TbrStyleStandard
            If (dwStyle And TBSTYLE_FLAT) = TBSTYLE_FLAT Then
                SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle And Not TBSTYLE_FLAT
                Call ReCreateButtons
                Call UserControl_Resize
            End If
        Case TbrStyleFlat
            If Not (dwStyle And TBSTYLE_FLAT) = TBSTYLE_FLAT Then
                SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle Or TBSTYLE_FLAT
                Call ReCreateButtons
                Call UserControl_Resize
            End If
    End Select
End If
UserControl.PropertyChanged "Style"
End Property

Public Property Get TextAlignment() As TbrTextAlignConstants
Attribute TextAlignment.VB_Description = "Returns/sets a value that determines whether button text is displayed below or to the right of the button image."
TextAlignment = PropTextAlignment
End Property

Public Property Let TextAlignment(ByVal Value As TbrTextAlignConstants)
Select Case Value
    Case TbrTextAlignBottom, TbrTextAlignRight
        PropTextAlignment = Value
    Case Else
        Err.Raise 380
End Select
If ToolBarHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
    Select Case PropTextAlignment
        Case TbrTextAlignBottom
            If (dwStyle And TBSTYLE_LIST) = TBSTYLE_LIST Then Call ReCreateToolBar
        Case TbrTextAlignRight
            If Not (dwStyle And TBSTYLE_LIST) = TBSTYLE_LIST Then Call ReCreateToolBar
    End Select
End If
UserControl.PropertyChanged "TextAlignment"
End Property

Public Property Get Orientation() As TbrOrientationConstants
Attribute Orientation.VB_Description = "Returns/sets the orientation for a non-aligned tool bar."
Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As TbrOrientationConstants)
Select Case Value
    Case TbrOrientationHorizontal, TbrOrientationVertical
        With UserControl
        Dim Align As Integer
        If ToolBarAlignable = True Then Align = .Extender.Align Else Align = vbAlignNone
        Select Case Align
            Case vbAlignTop, vbAlignBottom
                If Value <> TbrOrientationHorizontal Then
                    If ToolBarDesignMode = True Then
                        MsgBox "Orientation must be 0 - Horizontal when Align is 1 - AlignTop or 2 - AlignBottom", vbCritical + vbOKOnly
                        Exit Property
                    Else
                        Err.Raise Number:=383, Description:="Orientation must be 0 - Horizontal when Align is 1 - AlignTop or 2 - AlignBottom"
                    End If
                End If
            Case vbAlignLeft, vbAlignRight
                If Value <> TbrOrientationVertical Then
                    If ToolBarDesignMode = True Then
                        MsgBox "Orientation must be 1 - Vertical when Align is 3 - AlignLeft or 4 - AlignRight", vbCritical + vbOKOnly
                        Exit Property
                    Else
                        Err.Raise Number:=383, Description:="Orientation must be 1 - Vertical when Align is 3 - AlignLeft or 4 - AlignRight"
                    End If
                End If
        End Select
        If PropOrientation <> Value Then
            ToolBarResizeFrozen = True
            .Extender.Move .Extender.Left, .Extender.Top, .Extender.Height, .Extender.Width
            ToolBarResizeFrozen = False
        End If
        End With
        PropOrientation = Value
    Case Else
        Err.Raise 380
End Select
Call UserControl_Resize
UserControl.PropertyChanged "Orientation"
End Property

Public Property Get Divider() As Boolean
Attribute Divider.VB_Description = "Returns/sets a value that determines whether a two-pixel highlight being drawn at the top of the control or not."
Divider = PropDivider
End Property

Public Property Let Divider(ByVal Value As Boolean)
PropDivider = Value
If ToolBarHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
    If PropDivider = True Then
        If (dwStyle And CCS_NODIVIDER) = CCS_NODIVIDER Then
            SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle And Not CCS_NODIVIDER
            Call ComCtlsFrameChanged(ToolBarHandle)
        End If
    Else
        If Not (dwStyle And CCS_NODIVIDER) = CCS_NODIVIDER Then
            SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle Or CCS_NODIVIDER
            Call ComCtlsFrameChanged(ToolBarHandle)
        End If
    End If
    Call UserControl_Resize
End If
UserControl.PropertyChanged "Divider"
End Property

Public Property Get ShowTips() As Boolean
Attribute ShowTips.VB_Description = "Returns/sets a value that determines whether the tool tip text properties will be displayed or not."
ShowTips = PropShowTips
End Property

Public Property Let ShowTips(ByVal Value As Boolean)
PropShowTips = Value
If ToolBarHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
    If PropShowTips = True Then
        If Not (dwStyle And TBSTYLE_TOOLTIPS) = TBSTYLE_TOOLTIPS Then
            SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle Or TBSTYLE_TOOLTIPS
            If ToolBarToolTipHandle <> NULL_PTR Then
                SendMessage ToolBarHandle, TB_SETTOOLTIPS, ToolBarToolTipHandle, ByVal 0&
            Else
                ToolBarToolTipHandle = SendMessage(ToolBarHandle, TB_GETTOOLTIPS, 0, ByVal 0&)
            End If
        End If
    Else
        If (dwStyle And TBSTYLE_TOOLTIPS) = TBSTYLE_TOOLTIPS Then
            ToolBarToolTipHandle = SendMessage(ToolBarHandle, TB_GETTOOLTIPS, 0, ByVal 0&)
            SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle And Not TBSTYLE_TOOLTIPS
            SendMessage ToolBarHandle, TB_SETTOOLTIPS, 0, ByVal 0&
        End If
    End If
End If
UserControl.PropertyChanged "ShowTips"
End Property

Public Property Get Wrappable() As Boolean
Attribute Wrappable.VB_Description = "Returns/sets whether buttons can be wrapped or not."
Wrappable = PropWrappable
End Property

Public Property Let Wrappable(ByVal Value As Boolean)
PropWrappable = Value
If ToolBarHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
    If PropWrappable = True Then
        If Not (dwStyle And TBSTYLE_WRAPABLE) = TBSTYLE_WRAPABLE Then SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle Or TBSTYLE_WRAPABLE
    Else
        If (dwStyle And TBSTYLE_WRAPABLE) = TBSTYLE_WRAPABLE Then SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle And Not TBSTYLE_WRAPABLE
    End If
    Me.Refresh
    Call UserControl_Resize
End If
UserControl.PropertyChanged "Wrappable"
End Property

Public Property Get AllowCustomize() As Boolean
Attribute AllowCustomize.VB_Description = "Returns/sets a value which determines if users can customize the tool bar."
AllowCustomize = PropAllowCustomize
End Property

Public Property Let AllowCustomize(ByVal Value As Boolean)
PropAllowCustomize = Value
If ToolBarHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
    If PropAllowCustomize = True Then
        If Not (dwStyle And CCS_ADJUSTABLE) = CCS_ADJUSTABLE Then SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle Or CCS_ADJUSTABLE
    Else
        If (dwStyle And CCS_ADJUSTABLE) = CCS_ADJUSTABLE Then SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle And Not CCS_ADJUSTABLE
    End If
End If
UserControl.PropertyChanged "AllowCustomize"
End Property

Public Property Get AltDrag() As Boolean
Attribute AltDrag.VB_Description = "Returns/sets a value indicating if users can change a button's position by dragging it while holding down the ALT key instead of the SHIFT key. The allow customize property need to be set to true."
AltDrag = PropAltDrag
End Property

Public Property Let AltDrag(ByVal Value As Boolean)
PropAltDrag = Value
If ToolBarHandle <> NULL_PTR Then
    Dim dwStyle As Long
    dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
    If PropAltDrag = True Then
        If Not (dwStyle And TBSTYLE_ALTDRAG) = TBSTYLE_ALTDRAG Then SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle Or TBSTYLE_ALTDRAG
    Else
        If (dwStyle And TBSTYLE_ALTDRAG) = TBSTYLE_ALTDRAG Then SendMessage ToolBarHandle, TB_SETSTYLE, 0, ByVal dwStyle And Not TBSTYLE_ALTDRAG
    End If
End If
UserControl.PropertyChanged "AltDrag"
End Property

Public Property Get DoubleBuffer() As Boolean
Attribute DoubleBuffer.VB_Description = "Returns/sets a value that determines whether the control paints via double-buffering, which reduces flicker."
DoubleBuffer = PropDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal Value As Boolean)
PropDoubleBuffer = Value
UserControl.PropertyChanged "DoubleBuffer"
End Property

Public Property Get ButtonHeight() As Single
Attribute ButtonHeight.VB_Description = "Returns/sets the height of the buttons."
If ToolBarHandle <> NULL_PTR Then
    If SendMessage(ToolBarHandle, TB_BUTTONCOUNT, 0, ByVal 0&) > 0 Then
        ButtonHeight = UserControl.ScaleY(HiWord(CLng(SendMessage(ToolBarHandle, TB_GETBUTTONSIZE, 0, ByVal 0&))), vbPixels, vbContainerSize)
    Else
        ButtonHeight = UserControl.ScaleY(PropButtonHeight, vbPixels, vbContainerSize)
    End If
Else
    ButtonHeight = UserControl.ScaleY(PropButtonHeight, vbPixels, vbContainerSize)
End If
End Property

Public Property Let ButtonHeight(ByVal Value As Single)
If Value < 0 Then
    If ToolBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim IntValue As Integer
On Error Resume Next
IntValue = CInt(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
If Err.Number <> 0 Then IntValue = 0
On Error GoTo 0
PropButtonHeight = IntValue
If PropButtonHeight < (22 * PixelsPerDIP_Y()) Then PropButtonHeight = (22 * PixelsPerDIP_Y())
If ToolBarHandle <> NULL_PTR And ToolBarDesignMode = False Then SendMessage ToolBarHandle, TB_SETBUTTONSIZE, 0, ByVal MakeDWord(LoWord(CLng(SendMessage(ToolBarHandle, TB_GETBUTTONSIZE, 0, ByVal 0&))), PropButtonHeight)
Call UserControl_Resize
UserControl.PropertyChanged "ButtonHeight"
End Property

Public Property Get ButtonWidth() As Single
Attribute ButtonWidth.VB_Description = "Returns/sets the width of the buttons."
If ToolBarHandle <> NULL_PTR Then
    If SendMessage(ToolBarHandle, TB_BUTTONCOUNT, 0, ByVal 0&) > 0 Then
        ButtonWidth = UserControl.ScaleX(LoWord(CLng(SendMessage(ToolBarHandle, TB_GETBUTTONSIZE, 0, ByVal 0&))), vbPixels, vbContainerSize)
    Else
        ButtonWidth = UserControl.ScaleX(PropButtonWidth, vbPixels, vbContainerSize)
    End If
Else
    ButtonWidth = UserControl.ScaleX(PropButtonWidth, vbPixels, vbContainerSize)
End If
End Property

Public Property Let ButtonWidth(ByVal Value As Single)
If Value < 0 Then
    If ToolBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim IntValue As Integer
On Error Resume Next
IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If Err.Number <> 0 Then IntValue = 0
On Error GoTo 0
PropButtonWidth = IntValue
If PropButtonWidth < (24 * PixelsPerDIP_X()) Then PropButtonWidth = (24 * PixelsPerDIP_X())
If ToolBarHandle <> NULL_PTR And ToolBarDesignMode = False Then SendMessage ToolBarHandle, TB_SETBUTTONSIZE, 0, ByVal MakeDWord(PropButtonWidth, HiWord(CLng(SendMessage(ToolBarHandle, TB_GETBUTTONSIZE, 0, ByVal 0&))))
Call UserControl_Resize
UserControl.PropertyChanged "ButtonWidth"
End Property

Public Property Get MinButtonWidth() As Single
Attribute MinButtonWidth.VB_Description = "Returns/sets the minimum width of the buttons."
MinButtonWidth = UserControl.ScaleX(PropMinButtonWidth, vbPixels, vbContainerSize)
End Property

Public Property Let MinButtonWidth(ByVal Value As Single)
If Value < 0 Then
    If ToolBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim IntValue As Integer, ErrValue As Long
On Error Resume Next
IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
ErrValue = Err.Number
On Error GoTo 0
If IntValue >= 0 And ErrValue = 0 Then
    PropMinButtonWidth = IntValue
    If ToolBarHandle <> NULL_PTR Then
        SendMessage ToolBarHandle, TB_SETBUTTONWIDTH, 0, ByVal MakeDWord(PropMinButtonWidth, PropMaxButtonWidth)
        Call ReCreateButtons
    End If
    Call UserControl_Resize
Else
    If ToolBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
UserControl.PropertyChanged "MinButtonWidth"
End Property

Public Property Get MaxButtonWidth() As Single
Attribute MaxButtonWidth.VB_Description = "Returns/sets the maximum width of the buttons."
MaxButtonWidth = UserControl.ScaleX(PropMaxButtonWidth, vbPixels, vbContainerSize)
End Property

Public Property Let MaxButtonWidth(ByVal Value As Single)
If Value < 0 Then
    If ToolBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim IntValue As Integer, ErrValue As Long
On Error Resume Next
IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
ErrValue = Err.Number
On Error GoTo 0
If IntValue >= 0 And ErrValue = 0 Then
    PropMaxButtonWidth = IntValue
    If ToolBarHandle <> NULL_PTR Then
        SendMessage ToolBarHandle, TB_SETBUTTONWIDTH, 0, ByVal MakeDWord(PropMinButtonWidth, PropMaxButtonWidth)
        Call ReCreateButtons
    End If
    Call UserControl_Resize
Else
    If ToolBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
UserControl.PropertyChanged "MaxButtonWidth"
End Property

Public Property Get InsertMarkColor() As OLE_COLOR
Attribute InsertMarkColor.VB_Description = "Returns/sets the color of the insertion mark."
InsertMarkColor = PropInsertMarkColor
End Property

Public Property Let InsertMarkColor(ByVal Value As OLE_COLOR)
PropInsertMarkColor = Value
If ToolBarHandle <> NULL_PTR Then SendMessage ToolBarHandle, TB_SETINSERTMARKCOLOR, 0, ByVal WinColor(PropInsertMarkColor)
UserControl.PropertyChanged "InsertMarkColor"
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

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets whether hot tracking is enabled."
HotTracking = PropHotTracking
End Property

Public Property Let HotTracking(ByVal Value As Boolean)
PropHotTracking = Value
Me.Refresh
UserControl.PropertyChanged "HotTracking"
End Property

Public Property Get HideClippedButtons() As Boolean
Attribute HideClippedButtons.VB_Description = "Returns/sets a value that determines whether or not partially clipped buttons are hidden."
HideClippedButtons = PropHideClippedButtons
End Property

Public Property Let HideClippedButtons(ByVal Value As Boolean)
PropHideClippedButtons = Value
If ToolBarHandle <> NULL_PTR Then
    Dim dwExStyle As Long
    dwExStyle = CLng(SendMessage(ToolBarHandle, TB_GETEXTENDEDSTYLE, 0, ByVal 0&))
    If PropHideClippedButtons = True Then
        If Not (dwExStyle And TBSTYLE_EX_HIDECLIPPEDBUTTONS) = TBSTYLE_EX_HIDECLIPPEDBUTTONS Then SendMessage ToolBarHandle, TB_SETEXTENDEDSTYLE, 0, ByVal dwExStyle Or TBSTYLE_EX_HIDECLIPPEDBUTTONS
    Else
        If (dwExStyle And TBSTYLE_EX_HIDECLIPPEDBUTTONS) = TBSTYLE_EX_HIDECLIPPEDBUTTONS Then SendMessage ToolBarHandle, TB_SETEXTENDEDSTYLE, 0, ByVal dwExStyle And Not TBSTYLE_EX_HIDECLIPPEDBUTTONS
    End If
End If
UserControl.PropertyChanged "HideClippedButtons"
End Property

Public Property Get AnchorHot() As Boolean
Attribute AnchorHot.VB_Description = "Returns/sets a value indicating if the currently hot button will remain hot even if the user moves the mouse out of the control."
If ToolBarHandle <> NULL_PTR Then
    AnchorHot = CBool(SendMessage(ToolBarHandle, TB_GETANCHORHIGHLIGHT, 0, ByVal 0&) <> 0)
Else
    AnchorHot = PropAnchorHot
End If
End Property

Public Property Let AnchorHot(ByVal Value As Boolean)
PropAnchorHot = Value
If ToolBarHandle <> NULL_PTR Then SendMessage ToolBarHandle, TB_SETANCHORHIGHLIGHT, IIf(PropAnchorHot = True, 1, 0), ByVal 0&
UserControl.PropertyChanged "AnchorHot"
End Property

Public Property Get MaxTextRows() As Integer
Attribute MaxTextRows.VB_Description = "Returns/sets the maximum number of text rows displayed on a button. Only applicable if the text alignment property is set to bottom and the value of the max button width property is greater than 0."
MaxTextRows = PropMaxTextRows
End Property

Public Property Let MaxTextRows(ByVal Value As Integer)
If Value < 1 Then Err.Raise 380
If Value > 1 And PropTextAlignment = TbrTextAlignRight Then
    If ToolBarDesignMode = True Then
        MsgBox "MaxTextRows must be 1 when TextAlignment is 1 - TextAlignRight", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise Number:=383, Description:="MaxTextRows must be 1 when TextAlignment is 1 - TextAlignRight"
    End If
End If
PropMaxTextRows = Value
If ToolBarHandle <> NULL_PTR Then SendMessage ToolBarHandle, TB_SETMAXTEXTROWS, PropMaxTextRows, ByVal 0&
Call UserControl_Resize
UserControl.PropertyChanged "MaxTextRows"
End Property

Public Property Get Buttons() As TbrButtons
Attribute Buttons.VB_Description = "Returns a reference to a collection of the button objects."
If PropButtons Is Nothing Then
    Set PropButtons = New TbrButtons
    PropButtons.FInit Me
End If
Set Buttons = PropButtons
End Property

Friend Sub FButtonsAdd(ByVal Index As Long, ByVal NewButton As TbrButton, Optional ByVal Caption As String, Optional ByVal Style As TbrButtonStyleConstants, Optional ByVal ImageIndex As Long)
Dim TBB As TBBUTTON
With TBB
.IDCommand = NextButtonID()
NewButton.ID = .IDCommand
.iBitmap = ImageIndex - 1
.iString = StrPtr(Caption)
.fsState = TBSTATE_ENABLED
Select Case Style
    Case TbrButtonStyleDefault
        .fsStyle = BTNS_BUTTON
    Case TbrButtonStyleCheck
        .fsStyle = BTNS_CHECK
    Case TbrButtonStyleCheckGroup
        .fsStyle = BTNS_CHECKGROUP
    Case TbrButtonStyleSeparator
        .fsStyle = BTNS_SEP
        .iBitmap = 0
        .iString = NULL_PTR
    Case TbrButtonStyleDropDown
        .fsStyle = BTNS_DROPDOWN
    Case TbrButtonStyleWholeDropDown
        .fsStyle = BTNS_WHOLEDROPDOWN
    Case Else
        Err.Raise 380
End Select
.dwData = ObjPtr(NewButton)
End With
If ToolBarHandle <> NULL_PTR Then
    Call ResetCustomizeButtons
    If Index = 0 Then
        SendMessage ToolBarHandle, TB_ADDBUTTONS, 1, ByVal VarPtr(TBB)
    Else
        SendMessage ToolBarHandle, TB_INSERTBUTTON, Index - 1, ByVal VarPtr(TBB)
    End If
    Dim Size As Long
    Size = CLng(SendMessage(ToolBarHandle, TB_GETBUTTONSIZE, 0, ByVal 0&))
    PropButtonWidth = LoWord(Size)
    PropButtonHeight = HiWord(Size)
End If
Call UserControl_Resize
UserControl.PropertyChanged "InitButtons"
End Sub

Friend Sub FButtonsRemove(ByVal ID As Long)
If ToolBarHandle <> NULL_PTR Then
    Call ResetCustomizeButtons
    SendMessage ToolBarHandle, TB_DELETEBUTTON, SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&), ByVal 0&
    Call ReCreateButtons
End If
Call UserControl_Resize
UserControl.PropertyChanged "InitButtons"
End Sub

Friend Sub FButtonsClear()
If ToolBarHandle <> NULL_PTR Then
    Call ResetCustomizeButtons
    Do While SendMessage(ToolBarHandle, TB_DELETEBUTTON, 0, ByVal 0&) <> 0: Loop
    Me.Refresh
End If
Call UserControl_Resize
End Sub

Friend Sub FButtonRedraw(ByVal ID As Long)
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then
    Dim RC As RECT
    ' TB_GETITEMRECT fails for buttons whose state is set to TBSTATE_HIDDEN, thus no need to redraw then.
    If SendMessage(ToolBarHandle, TB_GETITEMRECT, SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&), ByVal VarPtr(RC)) <> 0 Then
        InvalidateRect ToolBarHandle, ByVal VarPtr(RC), 1
        UpdateWindow ToolBarHandle
        UserControl.Refresh
    End If
End If
End Sub

Friend Property Let FButtonCaption(ByVal ID As Long, ByVal Value As String)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Call ResetCustomizeButtons
        Dim NewButton As ShadowButtonStruct
        NewButton = GetShadowButton(ID)
        NewButton.Caption = Value
        Call ModifyButton(ID, NewButton)
        Call UserControl_Resize
    End If
End If
End Property

Friend Property Get FButtonStyle(ByVal ID As Long, ByVal ImageIndex As Long) As TbrButtonStyleConstants
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_STYLE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        If (.fsStyle And BTNS_WHOLEDROPDOWN) = BTNS_WHOLEDROPDOWN Then
            FButtonStyle = TbrButtonStyleWholeDropDown
        ElseIf (.fsStyle And BTNS_DROPDOWN) = BTNS_DROPDOWN Then
            FButtonStyle = TbrButtonStyleDropDown
        ElseIf (.fsStyle And BTNS_CHECKGROUP) = BTNS_CHECKGROUP Then
            FButtonStyle = TbrButtonStyleCheckGroup
        ElseIf (.fsStyle And BTNS_CHECK) = BTNS_CHECK Then
            FButtonStyle = TbrButtonStyleCheck
        ElseIf (.fsStyle And BTNS_SEP) = BTNS_SEP Then
            FButtonStyle = TbrButtonStyleSeparator
        ElseIf (.fsStyle And BTNS_BUTTON) = BTNS_BUTTON Then
            FButtonStyle = TbrButtonStyleDefault
        End If
        End With
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then
                If (.fsStyle And BTNS_WHOLEDROPDOWN) = BTNS_WHOLEDROPDOWN Then
                    FButtonStyle = TbrButtonStyleWholeDropDown
                ElseIf (.fsStyle And BTNS_DROPDOWN) = BTNS_DROPDOWN Then
                    FButtonStyle = TbrButtonStyleDropDown
                ElseIf (.fsStyle And BTNS_CHECKGROUP) = BTNS_CHECKGROUP Then
                    FButtonStyle = TbrButtonStyleCheckGroup
                ElseIf (.fsStyle And BTNS_CHECK) = BTNS_CHECK Then
                    FButtonStyle = TbrButtonStyleCheck
                ElseIf (.fsStyle And BTNS_SEP) = BTNS_SEP Then
                    FButtonStyle = TbrButtonStyleSeparator
                ElseIf (.fsStyle And BTNS_BUTTON) = BTNS_BUTTON Then
                    FButtonStyle = TbrButtonStyleDefault
                End If
            End If
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonStyle(ByVal ID As Long, ByVal ImageIndex As Long, ByVal Value As TbrButtonStyleConstants)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Call ResetCustomizeButtons
        Dim OldButton As ShadowButtonStruct, NewButton As ShadowButtonStruct
        OldButton = GetShadowButton(ID)
        With NewButton
        LSet .TBB = OldButton.TBB
        With .TBB
        If (.fsState And TBSTATE_CHECKED) = TBSTATE_CHECKED Then .fsState = .fsState And Not TBSTATE_CHECKED
        If (.fsState And TBSTATE_PRESSED) = TBSTATE_PRESSED Then .fsState = .fsState And Not TBSTATE_PRESSED
        If (.fsStyle And BTNS_SEP) = BTNS_SEP Then .fsStyle = .fsStyle And Not BTNS_SEP
        If (.fsStyle And BTNS_CHECK) = BTNS_CHECK Then .fsStyle = .fsStyle And Not BTNS_CHECK
        If (.fsStyle And BTNS_GROUP) = BTNS_GROUP Then .fsStyle = .fsStyle And Not BTNS_GROUP
        If (.fsStyle And BTNS_DROPDOWN) = BTNS_DROPDOWN Then .fsStyle = .fsStyle And Not BTNS_DROPDOWN
        If (.fsStyle And BTNS_WHOLEDROPDOWN) = BTNS_WHOLEDROPDOWN Then .fsStyle = .fsStyle And Not BTNS_WHOLEDROPDOWN
        .iBitmap = ImageIndex - 1
        Select Case Value
            Case TbrButtonStyleDefault
                .fsStyle = .fsStyle Or BTNS_BUTTON
            Case TbrButtonStyleCheck
                .fsStyle = .fsStyle Or BTNS_CHECK
            Case TbrButtonStyleCheckGroup
                .fsStyle = .fsStyle Or BTNS_CHECKGROUP
            Case TbrButtonStyleSeparator
                .fsStyle = .fsStyle Or BTNS_SEP
                .iBitmap = 0
                .iString = 0
            Case TbrButtonStyleDropDown
                .fsStyle = .fsStyle Or BTNS_DROPDOWN
            Case TbrButtonStyleWholeDropDown
                .fsStyle = .fsStyle Or BTNS_WHOLEDROPDOWN
            Case Else
                Err.Raise 380
        End Select
        End With
        .Caption = GetButtonText(ID)
        .CX = OldButton.CX
        End With
        Call ModifyButton(ID, NewButton)
        Call UserControl_Resize
    End If
End If
End Property

Friend Property Let FButtonImage(ByVal ID As Long, ByVal Value As Long)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_STYLE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        If (.fsStyle And BTNS_SEP) = 0 Then
            .dwMask = TBIF_IMAGE
            .iImage = Value - 1
            SendMessage ToolBarHandle, TB_SETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        End If
        End With
    End If
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then .iBitmap = Value - 1
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonValue(ByVal ID As Long) As TbrButtonValueConstants
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_STYLE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        If (.fsStyle And BTNS_CHECK) <> 0 Then
            If SendMessage(ToolBarHandle, TB_ISBUTTONCHECKED, ID, ByVal 0&) <> 0 Then
                FButtonValue = TbrButtonValuePressed
            Else
                FButtonValue = TbrButtonValueUnpressed
            End If
        Else
            If SendMessage(ToolBarHandle, TB_ISBUTTONPRESSED, ID, ByVal 0&) <> 0 Then
                FButtonValue = TbrButtonValuePressed
            Else
                FButtonValue = TbrButtonValueUnpressed
            End If
        End If
        End With
    Else
        FButtonValue = TbrButtonValueUnpressed
    End If
End If
End Property

Friend Property Let FButtonValue(ByVal ID As Long, ByVal Value As TbrButtonValueConstants)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_STYLE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        If (.fsStyle And BTNS_CHECK) <> 0 Then
            SendMessage ToolBarHandle, TB_CHECKBUTTON, ID, ByVal MakeDWord(IIf(Value = TbrButtonValuePressed, 1, 0), 0)
        Else
            SendMessage ToolBarHandle, TB_PRESSBUTTON, ID, ByVal MakeDWord(IIf(Value = TbrButtonValuePressed, 1, 0), 0)
        End If
        End With
    End If
End If
End Property

Friend Property Get FButtonEnabled(ByVal ID As Long) As Boolean
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        FButtonEnabled = CBool(SendMessage(ToolBarHandle, TB_ISBUTTONENABLED, ID, ByVal 0&) <> 0)
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then FButtonEnabled = CBool((.fsState And TBSTATE_ENABLED) = TBSTATE_ENABLED)
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonEnabled(ByVal ID As Long, ByVal Value As Boolean)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then SendMessage ToolBarHandle, TB_ENABLEBUTTON, ID, ByVal MakeDWord(IIf(Value = True, 1, 0), 0)
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i)
            If .TBB.IDCommand = ID Then
                If Value = True Then
                    If Not (.TBB.fsState And TBSTATE_ENABLED) = TBSTATE_ENABLED Then .TBB.fsState = .TBB.fsState Or TBSTATE_ENABLED
                Else
                    If (.TBB.fsState And TBSTATE_ENABLED) = TBSTATE_ENABLED Then .TBB.fsState = .TBB.fsState And Not TBSTATE_ENABLED
                End If
            End If
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonVisible(ByVal ID As Long) As Boolean
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        FButtonVisible = CBool(SendMessage(ToolBarHandle, TB_ISBUTTONHIDDEN, ID, ByVal 0&) = 0)
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then FButtonVisible = Not CBool((.fsState And TBSTATE_HIDDEN) = TBSTATE_HIDDEN)
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonVisible(ByVal ID As Long, ByVal Value As Boolean)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then SendMessage ToolBarHandle, TB_HIDEBUTTON, ID, ByVal MakeDWord(IIf(Value = True, 0, 1), 0)
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i)
            If .TBB.IDCommand = ID Then
                If Value = True Then
                    If (.TBB.fsState And TBSTATE_HIDDEN) = TBSTATE_HIDDEN Then .TBB.fsState = .TBB.fsState And Not TBSTATE_HIDDEN
                Else
                    If Not (.TBB.fsState And TBSTATE_HIDDEN) = TBSTATE_HIDDEN Then .TBB.fsState = .TBB.fsState Or TBSTATE_HIDDEN
                End If
            End If
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonMixedState(ByVal ID As Long) As Boolean
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        FButtonMixedState = CBool(SendMessage(ToolBarHandle, TB_ISBUTTONINDETERMINATE, ID, ByVal 0&) <> 0)
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then FButtonMixedState = CBool((.fsState And TBSTATE_INDETERMINATE) = TBSTATE_INDETERMINATE)
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonMixedState(ByVal ID As Long, ByVal Value As Boolean)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then SendMessage ToolBarHandle, TB_INDETERMINATE, ID, ByVal MakeDWord(IIf(Value = True, 1, 0), 0)
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i)
            If .TBB.IDCommand = ID Then
                If Value = True Then
                    If Not (.TBB.fsState And TBSTATE_INDETERMINATE) = TBSTATE_INDETERMINATE Then .TBB.fsState = .TBB.fsState Or TBSTATE_INDETERMINATE
                Else
                    If (.TBB.fsState And TBSTATE_INDETERMINATE) = TBSTATE_INDETERMINATE Then .TBB.fsState = .TBB.fsState And Not TBSTATE_INDETERMINATE
                End If
            End If
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonHighLighted(ByVal ID As Long) As Boolean
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        FButtonHighLighted = CBool(SendMessage(ToolBarHandle, TB_ISBUTTONHIGHLIGHTED, ID, ByVal 0&) <> 0)
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then FButtonHighLighted = CBool((.fsState And TBSTATE_MARKED) = TBSTATE_MARKED)
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonHighLighted(ByVal ID As Long, ByVal Value As Boolean)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then SendMessage ToolBarHandle, TB_MARKBUTTON, ID, ByVal MakeDWord(IIf(Value = True, 1, 0), 0)
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i)
            If .TBB.IDCommand = ID Then
                If Value = True Then
                    If Not (.TBB.fsState And TBSTATE_MARKED) = TBSTATE_MARKED Then .TBB.fsState = .TBB.fsState Or TBSTATE_MARKED
                Else
                    If (.TBB.fsState And TBSTATE_MARKED) = TBSTATE_MARKED Then .TBB.fsState = .TBB.fsState And Not TBSTATE_MARKED
                End If
            End If
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonNoImage(ByVal ID As Long, ByVal ImageIndex As Long) As Boolean
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_STYLE Or TBIF_IMAGE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        FButtonNoImage = CBool(.iImage = I_IMAGENONE And (.fsStyle And BTNS_SEP) = 0)
        End With
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then FButtonNoImage = CBool(.iBitmap = I_IMAGENONE And (.fsStyle And BTNS_SEP) = 0)
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonNoImage(ByVal ID As Long, ByVal ImageIndex As Long, ByVal Value As Boolean)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim NewButton As ShadowButtonStruct
        NewButton = GetShadowButton(ID)
        If (NewButton.TBB.fsStyle And BTNS_SEP) = 0 Then
            With NewButton
            .TBB.iBitmap = IIf(Value, I_IMAGENONE, ImageIndex - 1)
            .Caption = GetButtonText(ID)
            End With
            Call ModifyButton(ID, NewButton)
            Call UserControl_Resize
        End If
    End If
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then
                If (.fsStyle And BTNS_SEP) = 0 Then
                    If Value = True Then
                        .iBitmap = I_IMAGENONE
                    Else
                        .iBitmap = ImageIndex - 1
                    End If
                End If
            End If
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonNoPrefix(ByVal ID As Long) As Boolean
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_STYLE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        FButtonNoPrefix = CBool((.fsStyle And BTNS_NOPREFIX) = BTNS_NOPREFIX)
        End With
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then FButtonNoPrefix = CBool((.fsStyle And BTNS_NOPREFIX) = BTNS_NOPREFIX)
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonNoPrefix(ByVal ID As Long, ByVal Value As Boolean)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim NewButton As ShadowButtonStruct
        With NewButton
        NewButton = GetShadowButton(ID)
        If Value = True Then
            If Not (.TBB.fsStyle And BTNS_NOPREFIX) = BTNS_NOPREFIX Then .TBB.fsStyle = .TBB.fsStyle Or BTNS_NOPREFIX
        Else
            If (.TBB.fsStyle And BTNS_NOPREFIX) = BTNS_NOPREFIX Then .TBB.fsStyle = .TBB.fsStyle And Not BTNS_NOPREFIX
        End If
        .Caption = GetButtonText(ID)
        End With
        Call ModifyButton(ID, NewButton)
        Call UserControl_Resize
    End If
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i)
            If .TBB.IDCommand = ID Then
                If Value = True Then
                    If Not (.TBB.fsStyle And BTNS_NOPREFIX) = BTNS_NOPREFIX Then .TBB.fsStyle = .TBB.fsStyle Or BTNS_NOPREFIX
                Else
                    If (.TBB.fsStyle And BTNS_NOPREFIX) = BTNS_NOPREFIX Then .TBB.fsStyle = .TBB.fsStyle And Not BTNS_NOPREFIX
                End If
            End If
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonAutoSize(ByVal ID As Long) As Boolean
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_STYLE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        FButtonAutoSize = CBool((.fsStyle And BTNS_AUTOSIZE) = BTNS_AUTOSIZE)
        End With
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then FButtonAutoSize = CBool((.fsStyle And BTNS_AUTOSIZE) = BTNS_AUTOSIZE)
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonAutoSize(ByVal ID As Long, ByVal Value As Boolean)
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim NewButton As ShadowButtonStruct
        NewButton = GetShadowButton(ID)
        With NewButton
        If Value = True Then
            If Not (.TBB.fsStyle And BTNS_AUTOSIZE) = BTNS_AUTOSIZE Then .TBB.fsStyle = .TBB.fsStyle Or BTNS_AUTOSIZE
        Else
            If (.TBB.fsStyle And BTNS_AUTOSIZE) = BTNS_AUTOSIZE Then .TBB.fsStyle = .TBB.fsStyle And Not BTNS_AUTOSIZE
        End If
        .Caption = GetButtonText(ID)
        .CX = 0
        End With
        Call ModifyButton(ID, NewButton)
        Call UserControl_Resize
    End If
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i)
            If .TBB.IDCommand = ID Then
                If Value = True Then
                    If Not (.TBB.fsStyle And BTNS_AUTOSIZE) = BTNS_AUTOSIZE Then .TBB.fsStyle = .TBB.fsStyle Or BTNS_AUTOSIZE
                Else
                    If (.TBB.fsStyle And BTNS_AUTOSIZE) = BTNS_AUTOSIZE Then .TBB.fsStyle = .TBB.fsStyle And Not BTNS_AUTOSIZE
                End If
            End If
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonCustomWidth(ByVal ID As Long) As Single
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_SIZE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        FButtonCustomWidth = UserControl.ScaleX(.CX, vbPixels, vbContainerSize)
        End With
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i)
            If .TBB.IDCommand = ID Then FButtonCustomWidth = .CX
            End With
        Next i
    End If
End If
End Property

Friend Property Let FButtonCustomWidth(ByVal ID As Long, ByVal Value As Single)
If Value < 0 Then
    If ToolBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim TBBI As TBBUTTONINFO
        With TBBI
        .cbSize = LenB(TBBI)
        .dwMask = TBIF_STYLE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        If (.fsStyle And BTNS_AUTOSIZE) = 0 Then
            .dwMask = TBIF_SIZE
            .CX = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
        Else
            .CX = 0
        End If
        SendMessage ToolBarHandle, TB_SETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        SendMessage ToolBarHandle, TB_AUTOSIZE, 0, ByVal 0&
        Me.Refresh
        End With
    End If
    If ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i)
            If .TBB.IDCommand = ID Then .CX = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
            End With
        Next i
    End If
End If
End Property

Friend Property Get FButtonPosition(ByVal ID As Long) As Long
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then FButtonPosition = CLng(SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&)) + 1
End Property

Friend Property Let FButtonPosition(ByVal ID As Long, ByVal Value As Long)
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then If SendMessage(ToolBarHandle, TB_MOVEBUTTON, SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&), ByVal Value - 1&) = 0 Then Err.Raise 380
End Property

Friend Property Get FButtonHot(ByVal ID As Long) As Boolean
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then
    Dim Index As Long
    Index = CLng(SendMessage(ToolBarHandle, TB_GETHOTITEM, 0, ByVal 0&))
    If Index > -1 Then FButtonHot = CBool(ID = GetButtonID(Index + 1))
End If
End Property

Friend Property Let FButtonHot(ByVal ID As Long, ByVal Value As Boolean)
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then
    Dim Index As Long
    Index = CLng(SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&)) + 1
    If Index > 0 Then
        If Value = True Then
            SendMessage ToolBarHandle, TB_SETHOTITEM, Index - 1, ByVal 0&
        Else
            If SendMessage(ToolBarHandle, TB_GETHOTITEM, 0, ByVal 0&) = (Index - 1) Then SendMessage ToolBarHandle, TB_SETHOTITEM, -1, ByVal 0&
        End If
    End If
End If
End Property

Friend Property Get FButtonWidth(ByVal ID As Long) As Single
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then
    Dim RC As RECT
    SendMessage ToolBarHandle, TB_GETITEMRECT, SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&), ByVal VarPtr(RC)
    FButtonWidth = UserControl.ScaleX((RC.Right - RC.Left), vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FButtonHeight(ByVal ID As Long) As Single
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then
    Dim RC As RECT
    SendMessage ToolBarHandle, TB_GETITEMRECT, SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&), ByVal VarPtr(RC)
    FButtonHeight = UserControl.ScaleX((RC.Bottom - RC.Top), vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FButtonLeft(ByVal ID As Long) As Single
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then
    Dim RC As RECT
    SendMessage ToolBarHandle, TB_GETITEMRECT, SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&), ByVal VarPtr(RC)
    FButtonLeft = UserControl.ScaleX(RC.Left, vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FButtonTop(ByVal ID As Long) As Single
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then
    Dim RC As RECT
    SendMessage ToolBarHandle, TB_GETITEMRECT, SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&), ByVal VarPtr(RC)
    FButtonTop = UserControl.ScaleX(RC.Top, vbPixels, vbContainerSize)
End If
End Property

Friend Property Get FButtonMenuParent(ByVal ID As Long) As TbrButton
If ToolBarHandle <> NULL_PTR Then
    If IsButtonAvailable(ID) = True Then
        Dim Ptr As LongPtr
        Ptr = GetButtonPtr(ID)
        If Ptr <> NULL_PTR Then Set FButtonMenuParent = PtrToObj(Ptr)
    ElseIf ToolBarCustomizeButtonsCount > 0 Then
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount
            With ToolBarCustomizeButtons(i).TBB
            If .IDCommand = ID Then If .dwData <> 0 Then Set FButtonMenuParent = PtrToObj(.dwData)
            End With
        Next i
    End If
End If
End Property

Private Sub CreateToolBar()
If ToolBarHandle <> NULL_PTR Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or CCS_NORESIZE
Dim Align As Integer
If ToolBarAlignable = True Then Align = Extender.Align Else Align = vbAlignNone
Select Case Align
    Case vbAlignNone
        dwStyle = dwStyle Or CCS_TOP
        If PropOrientation = TbrOrientationVertical Then dwStyle = dwStyle Or CCS_VERT
    Case vbAlignTop, vbAlignBottom
        dwStyle = dwStyle Or CCS_TOP
    Case vbAlignLeft, vbAlignRight
        dwStyle = dwStyle Or CCS_TOP Or CCS_VERT
End Select
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If PropStyle = TbrStyleFlat Then dwStyle = dwStyle Or TBSTYLE_FLAT
If PropTextAlignment = TbrTextAlignRight Then
    dwStyle = dwStyle Or TBSTYLE_LIST
    PropMaxTextRows = 1
End If
If PropDivider = False Then dwStyle = dwStyle Or CCS_NODIVIDER
If PropShowTips = True Then dwStyle = dwStyle Or TBSTYLE_TOOLTIPS
If PropWrappable = True Then dwStyle = dwStyle Or TBSTYLE_WRAPABLE
If PropAllowCustomize = True Then dwStyle = dwStyle Or CCS_ADJUSTABLE
If PropAltDrag = True Then dwStyle = dwStyle Or TBSTYLE_ALTDRAG
If ToolBarDesignMode = True Then
    dwStyle = dwStyle Or TBSTYLE_TRANSPARENT
    dwExStyle = dwExStyle Or WS_EX_TRANSPARENT
End If
ToolBarHandle = CreateWindowEx(dwExStyle, StrPtr("ToolbarWindow32"), NULL_PTR, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
If ToolBarHandle <> NULL_PTR Then
    Call ComCtlsShowAllUIStates(ToolBarHandle)
    Dim TBB As TBBUTTON
    SendMessage ToolBarHandle, TB_BUTTONSTRUCTSIZE, LenB(TBB), ByVal 0&
    SendMessage ToolBarHandle, TB_SETUNICODEFORMAT, 1, ByVal 0&
    SendMessage ToolBarHandle, TB_SETBUTTONSIZE, 0, ByVal MakeDWord(PropButtonWidth, PropButtonHeight)
    ' By default Windows creates an image list with 16x15 pixels but computes the button size as if the images were 16x16 pixels.
    ToolBarDefaultImageSize = MakeDWord(16, 16)
    SendMessage ToolBarHandle, TB_SETBITMAPSIZE, 0, ByVal ToolBarDefaultImageSize
    ToolBarImageSize = ToolBarDefaultImageSize
    SendMessage ToolBarHandle, TB_SETEXTENDEDSTYLE, 0, ByVal TBSTYLE_EX_DRAWDDARROWS
    SendMessage ToolBarHandle, TB_SETBUTTONWIDTH, 0, ByVal MakeDWord(PropMinButtonWidth, PropMaxButtonWidth)
    SendMessage ToolBarHandle, TB_SETMAXTEXTROWS, PropMaxTextRows, ByVal 0&
    If PropRightToLeft = True And PropRightToLeftLayout = False Then
        Dim DrawFlags As Long
        DrawFlags = CLng(SendMessage(ToolBarHandle, TB_SETDRAWTEXTFLAGS, 0, ByVal 0&))
        If Not (DrawFlags And DT_RTLREADING) = DT_RTLREADING Then DrawFlags = DrawFlags Or DT_RTLREADING
        DrawFlags = CLng(SendMessage(ToolBarHandle, TB_SETDRAWTEXTFLAGS, DrawFlags, ByVal DrawFlags))
    End If
End If
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.InsertMarkColor = PropInsertMarkColor
Me.HideClippedButtons = PropHideClippedButtons
Me.AnchorHot = PropAnchorHot
If ToolBarDesignMode = False Then
    If ToolBarHandle <> NULL_PTR Then
        If ToolBarBackColorBrush = NULL_PTR Then ToolBarBackColorBrush = CreateSolidBrush(WinColor(PropBackColor))
        Call ComCtlsSetSubclass(ToolBarHandle, Me, 1)
    End If
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
Else
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
End If
UserControl.BackColor = PropBackColor
End Sub

Private Sub ReCreateToolBar()
Dim Locked As Boolean
With Me
Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
Dim ReInitButtonsCount As Long
Dim ReInitButtons() As InitButtonStruct
ReInitButtonsCount = .Buttons.Count
If ReInitButtonsCount > 0 Then
    ReDim ReInitButtons(1 To ReInitButtonsCount) As InitButtonStruct
    Dim i As Long, ii As Long
    For i = 1 To ReInitButtonsCount
        With .Buttons(i)
        ReInitButtons(i).Key = .Key
        ReInitButtons(i).Tag = .Tag
        ReInitButtons(i).Caption = .Caption
        ReInitButtons(i).Style = .Style
        ReInitButtons(i).Image = .Image
        ReInitButtons(i).ImageIndex = .ImageIndex
        ReInitButtons(i).ToolTipText = .ToolTipText
        ReInitButtons(i).Description = .Description
        ReInitButtons(i).Value = .Value
        ReInitButtons(i).ForeColor = .ForeColor
        ReInitButtons(i).Enabled = .Enabled
        ReInitButtons(i).Visible = .Visible
        ReInitButtons(i).MixedState = .MixedState
        ReInitButtons(i).NoImage = .NoImage
        ReInitButtons(i).NoPrefix = .NoPrefix
        ReInitButtons(i).AutoSize = .AutoSize
        ReInitButtons(i).CustomWidth = UserControl.ScaleX(.CustomWidth, vbContainerSize, vbPixels)
        ReInitButtons(i).ButtonMenusCount = .ButtonMenus.Count
        If ReInitButtons(i).ButtonMenusCount > 0 Then
            ReDim ReInitButtons(i).ButtonMenus(1 To ReInitButtons(i).ButtonMenusCount)
            For ii = 1 To ReInitButtons(i).ButtonMenusCount
                With .ButtonMenus(ii)
                ReInitButtons(i).ButtonMenus(ii).Key = .Key
                ReInitButtons(i).ButtonMenus(ii).Tag = .Tag
                ReInitButtons(i).ButtonMenus(ii).Text = .Text
                ReInitButtons(i).ButtonMenus(ii).Enabled = .Enabled
                ReInitButtons(i).ButtonMenus(ii).Visible = .Visible
                ReInitButtons(i).ButtonMenus(ii).Checked = .Checked
                ReInitButtons(i).ButtonMenus(ii).Separator = .Separator
                End With
            Next ii
        End If
        End With
    Next i
End If
.Buttons.Clear
Call DestroyToolBar
Call CreateToolBar
Call UserControl_Resize
If ToolBarDesignMode = False Then
    If Not PropImageListControl Is Nothing Then Set .ImageList = PropImageListControl
    If Not PropDisabledImageListControl Is Nothing Then Set .DisabledImageList = PropDisabledImageListControl
    If Not PropHotImageListControl Is Nothing Then Set .HotImageList = PropHotImageListControl
    If Not PropPressedImageListControl Is Nothing Then Set .PressedImageList = PropPressedImageListControl
Else
    If Not PropImageListName = "(None)" Then .ImageList = PropImageListName
    If Not PropDisabledImageListName = "(None)" Then .DisabledImageList = PropDisabledImageListName
    If Not PropHotImageListName = "(None)" Then .HotImageList = PropHotImageListName
    If Not PropPressedImageListName = "(None)" Then .PressedImageList = PropPressedImageListName
End If
If ReInitButtonsCount > 0 Then
    For i = 1 To ReInitButtonsCount
        With .Buttons.Add(i, ReInitButtons(i).Key, ReInitButtons(i).Caption, ReInitButtons(i).Style, ReInitButtons(i).ImageIndex)
        .FInit Me, ReInitButtons(i).Key, ReInitButtons(i).Caption, ReInitButtons(i).Image, ReInitButtons(i).ImageIndex
        .Tag = ReInitButtons(i).Tag
        .ToolTipText = ReInitButtons(i).ToolTipText
        .Description = ReInitButtons(i).Description
        If ReInitButtons(i).Value = TbrButtonValuePressed Then .Value = TbrButtonValuePressed
        .ForeColor = ReInitButtons(i).ForeColor
        If ReInitButtons(i).Enabled = False Then .Enabled = False
        If ReInitButtons(i).Visible = False Then .Visible = False
        If ReInitButtons(i).MixedState = True Then .MixedState = True
        If ReInitButtons(i).NoImage = True Then .NoImage = True
        If ReInitButtons(i).NoPrefix = True Then .NoPrefix = True
        If ReInitButtons(i).AutoSize = True Then .AutoSize = True
        If ReInitButtons(i).CustomWidth > 0 Then .CustomWidth = UserControl.ScaleX(ReInitButtons(i).CustomWidth, vbPixels, vbContainerSize)
        If ReInitButtons(i).ButtonMenusCount > 0 Then
            For ii = 1 To ReInitButtons(i).ButtonMenusCount
                With .ButtonMenus.Add(ii, ReInitButtons(i).ButtonMenus(ii).Key, ReInitButtons(i).ButtonMenus(ii).Text)
                If ReInitButtons(i).ButtonMenus(ii).Enabled = False Then .Enabled = False
                If ReInitButtons(i).ButtonMenus(ii).Visible = False Then .Visible = False
                If ReInitButtons(i).ButtonMenus(ii).Checked = True Then .Checked = True
                If ReInitButtons(i).ButtonMenus(ii).Separator = True Then .Separator = True
                End With
            Next ii
        End If
        End With
    Next i
End If
If Locked = True Then LockWindowUpdate NULL_PTR
.Refresh
End With
End Sub

Private Sub DestroyToolBar()
If ToolBarHandle = NULL_PTR Then Exit Sub
Call ComCtlsRemoveSubclass(ToolBarHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow ToolBarHandle, SW_HIDE
SetParent ToolBarHandle, NULL_PTR
DestroyWindow ToolBarHandle
ToolBarHandle = NULL_PTR
ToolBarToolTipHandle = NULL_PTR
If ToolBarFontHandle <> NULL_PTR Then
    DeleteObject ToolBarFontHandle
    ToolBarFontHandle = NULL_PTR
End If
If ToolBarTransparentBrush <> NULL_PTR Then
    DeleteObject ToolBarTransparentBrush
    ToolBarTransparentBrush = NULL_PTR
End If
If ToolBarBackColorBrush <> NULL_PTR Then
    DeleteObject ToolBarBackColorBrush
    ToolBarBackColorBrush = NULL_PTR
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
If ToolBarTransparentBrush <> NULL_PTR Then
    DeleteObject ToolBarTransparentBrush
    ToolBarTransparentBrush = NULL_PTR
End If
UserControl.Refresh
RedrawWindow UserControl.hWnd, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get ContainedControls() As VBRUN.ContainedControls
Attribute ContainedControls.VB_Description = "Returns a collection that allows access to the controls contained within the control that were added to the control by the developer who uses the control."
Set ContainedControls = UserControl.ContainedControls
End Property

Public Sub Customize()
Attribute Customize.VB_Description = "Invokes the customization dialog."
If ToolBarHandle <> NULL_PTR Then SendMessage ToolBarHandle, TB_CUSTOMIZE, 0, ByVal 0&
End Sub

Public Sub GetIdealSize(ByRef Width As Single, ByRef Height As Single)
Attribute GetIdealSize.VB_Description = "Gets the ideal size to display all buttons."
If ToolBarHandle <> NULL_PTR Then
    Width = 0
    Height = 0
    Dim dwStyle As Long, Count As Long, Size As SIZEAPI
    dwStyle = CLng(SendMessage(ToolBarHandle, TB_GETSTYLE, 0, ByVal 0&))
    Count = CLng(SendMessage(ToolBarHandle, TB_BUTTONCOUNT, 0, ByVal 0&))
    With UserControl
    If Count > 0 Then
        Dim i As Long, RC As RECT
        For i = 0 To Count
            If SendMessage(ToolBarHandle, TB_GETITEMRECT, i, ByVal VarPtr(RC)) <> 0 Then
                If RC.Right > Size.CX Then Size.CX = RC.Right
                If RC.Bottom > Size.CY Then Size.CY = RC.Bottom
            End If
        Next i
    Else
        Size.CX = PropButtonWidth
        Size.CY = PropButtonHeight
    End If
    If Not (dwStyle And CCS_NODIVIDER) = CCS_NODIVIDER Then
        ' The divider line is a two-pixel highlight.
        Size.CY = Size.CY + 2
    End If
    Width = .ScaleX(Size.CX, vbPixels, vbContainerSize)
    Height = .ScaleY(Size.CY, vbPixels, vbContainerSize)
    End With
End If
End Sub

Public Sub SaveToolBar(ByVal Key As String, ByVal SubKey As String, ByVal Value As String)
Attribute SaveToolBar.VB_Description = "Saves a tool bar configuration."
If ToolBarHandle <> NULL_PTR Then
    Const HKEY_CURRENT_USER As Long = &H80000001
    Dim TBSP As TBSAVEPARAMS, Buffer As String
    Buffer = Key & "\" & SubKey
    With TBSP
    .hKey = HKEY_CURRENT_USER
    .pszSubKey = StrPtr(Buffer)
    .pszValueName = StrPtr(Value)
    End With
    SendMessage ToolBarHandle, TB_SAVERESTORE, 1, ByVal VarPtr(TBSP)
End If
End Sub

Public Sub RestoreToolBar(ByVal Key As String, ByVal SubKey As String, ByVal Value As String)
Attribute RestoreToolBar.VB_Description = "Restores a tool bar to its original state after being customized."
If ToolBarHandle <> NULL_PTR Then
    Const HKEY_CURRENT_USER As Long = &H80000001
    Dim TBSP As TBSAVEPARAMS, Buffer As String
    Buffer = Key & "\" & SubKey
    With TBSP
    .hKey = HKEY_CURRENT_USER
    .pszSubKey = StrPtr(Buffer)
    .pszValueName = StrPtr(Value)
    End With
    SendMessage ToolBarHandle, TB_SAVERESTORE, 0, ByVal VarPtr(TBSP)
End If
End Sub

Public Function ContainerKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer) As TbrButton
Attribute ContainerKeyDown.VB_Description = "Provides accelerator key access by forwarding the key down events of the container. The key preview property need to be set to true by a form container."
If ToolBarHandle = NULL_PTR Or (Shift <> vbAltMask And Shift <> (vbAltMask Or vbShiftMask)) Then Exit Function
If IsWindowEnabled(ToolBarHandle) = 0 Then Exit Function
Dim ID As Long, Accel As Integer, Count As Long, TBBI As TBBUTTONINFO
Count = CLng(SendMessage(ToolBarHandle, TB_BUTTONCOUNT, 0, ByVal 0&))
With TBBI
.cbSize = LenB(TBBI)
If Count > 0 Then
    Dim i As Long
    .dwMask = TBIF_STATE Or TBIF_STYLE
    For i = 1 To Count
        ID = GetButtonID(i)
        If ID > 0 Then
            SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
            If (.fsState And TBSTATE_ENABLED) <> 0 And (.fsStyle And BTNS_NOPREFIX) = 0 Then
                Accel = AccelCharCode(GetButtonText(ID))
                If (VkKeyScan(Accel) And &HFF&) = (KeyCode And &HFF&) Then Exit For Else ID = 0
            Else
                ID = 0
            End If
        End If
    Next i
End If
If ID > 0 Then
    .dwMask = TBIF_LPARAM Or TBIF_STYLE
    SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
    If .lParam <> 0 Then
        Dim Button As TbrButton
        Set Button = PtrToObj(.lParam)
        If (.fsStyle And BTNS_CHECK) <> 0 Then
            If (.fsStyle And BTNS_GROUP) <> 0 Then
                If SendMessage(ToolBarHandle, TB_ISBUTTONCHECKED, ID, ByVal 0&) = 0 Then
                    KeyCode = 0
                    SendMessage ToolBarHandle, TB_CHECKBUTTON, ID, ByVal 1&
                    RaiseEvent ButtonClick(Button)
                End If
            Else
                KeyCode = 0
                If SendMessage(ToolBarHandle, TB_ISBUTTONCHECKED, ID, ByVal 0&) = 0 Then
                    SendMessage ToolBarHandle, TB_CHECKBUTTON, ID, ByVal 1&
                Else
                    SendMessage ToolBarHandle, TB_CHECKBUTTON, ID, ByVal 0&
                End If
                RaiseEvent ButtonClick(Button)
            End If
        Else
            KeyCode = 0
            If (.fsStyle And BTNS_WHOLEDROPDOWN) = 0 Then
                SendMessage ToolBarHandle, TB_PRESSBUTTON, ID, ByVal 1&
                UpdateWindow ToolBarHandle
                Sleep 50
                SendMessage ToolBarHandle, TB_PRESSBUTTON, ID, ByVal 0&
                RaiseEvent ButtonClick(Button)
            Else
                If ToolBarPopupMenuHandle = NULL_PTR Then
                    SendMessage ToolBarHandle, TB_PRESSBUTTON, ID, ByVal 1&
                    RaiseEvent ButtonDropDown(Button)
                    Dim MenuItem As Long
                    MenuItem = ShowButtonMenuItems(Button, True)
                    If MenuItem >= 1 And MenuItem <= Button.ButtonMenus.Count Then RaiseEvent ButtonMenuClick(Button.ButtonMenus(MenuItem))
                    SendMessage ToolBarHandle, TB_PRESSBUTTON, ID, ByVal 0&
                Else
                    SendMessage ToolBarHandle, WM_CANCELMODE, 0, ByVal 0&
                End If
            End If
        End If
        Set ContainerKeyDown = Button
    End If
End If
End With
End Function

Public Function FindMnemonic(ByVal CharCode As Long) As TbrButton
Attribute FindMnemonic.VB_Description = "Returns a reference to the button object with an matching mnemonic character."
If ToolBarHandle <> NULL_PTR Then
    ' TB_MAPACCELERATOR matches either the mnemonic character or the first character in a button item.
    ' This behavior is undocumented and unwanted.
    ' The fix is to use the ID only when TB_MAPACCELERATOR returns a nonzero value.
    Dim ID As Long
    If SendMessage(ToolBarHandle, TB_MAPACCELERATOR, CharCode, ByVal VarPtr(ID)) <> 0 Then
        If IsButtonAvailable(ID) = True Then
            Dim Ptr As LongPtr
            Ptr = GetButtonPtr(ID)
            If Ptr <> NULL_PTR Then Set FindMnemonic = PtrToObj(Ptr)
        End If
    End If
End If
End Function

Public Function HitTest(ByVal X As Single, ByVal Y As Single) As TbrButton
Attribute HitTest.VB_Description = "Returns a reference to the button object located at the coordinates of X and Y."
If ToolBarHandle <> NULL_PTR Then
    Dim P As POINTAPI, Index As Long
    P.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    P.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    Index = CLng(SendMessage(ToolBarHandle, TB_HITTEST, 0, ByVal VarPtr(P))) + 1
    If Index > 0 Then
        Dim ID As Long
        ID = GetButtonID(Index)
        If IsButtonAvailable(ID) = True Then
            Dim Ptr As LongPtr
            Ptr = GetButtonPtr(ID)
            If Ptr <> NULL_PTR Then Set HitTest = PtrToObj(Ptr)
        End If
    End If
End If
End Function

Public Function HitTestInsertMark(ByVal X As Single, ByVal Y As Single, Optional ByRef After As Boolean) As TbrButton
Attribute HitTestInsertMark.VB_Description = "Returns a reference to the button object located at the coordinates of X and Y and retrieves a value that determines where the insertion point should appear."
If ToolBarHandle <> NULL_PTR Then
    Dim P As POINTAPI, TBIM As TBINSERTMARK
    P.X = CLng(UserControl.ScaleX(X, vbContainerPosition, vbPixels))
    P.Y = CLng(UserControl.ScaleY(Y, vbContainerPosition, vbPixels))
    With TBIM
    SendMessage ToolBarHandle, TB_INSERTMARKHITTEST, VarPtr(P), ByVal VarPtr(TBIM)
    If .iButton > -1 And (.dwFlags And TBIMHT_BACKGROUND) = 0 Then
        Dim ID As Long
        ID = GetButtonID(.iButton + 1)
        If IsButtonAvailable(ID) = True Then
            Dim Ptr As LongPtr
            Ptr = GetButtonPtr(ID)
            If Ptr <> NULL_PTR Then Set HitTestInsertMark = PtrToObj(Ptr)
        End If
    End If
    After = CBool((.dwFlags And TBIMHT_AFTER) <> 0 And (.dwFlags And TBIMHT_BACKGROUND) = 0)
    End With
End If
End Function

Public Property Get InsertMark(Optional ByRef After As Boolean) As TbrButton
Attribute InsertMark.VB_Description = "Returns/sets a reference to a button where an insertion mark is positioned."
Attribute InsertMark.VB_MemberFlags = "400"
If ToolBarHandle <> NULL_PTR Then
    Dim TBIM As TBINSERTMARK
    With TBIM
    SendMessage ToolBarHandle, TB_GETINSERTMARK, 0, ByVal VarPtr(TBIM)
    If .iButton > -1 Then
        Dim ID As Long
        ID = GetButtonID(.iButton + 1)
        If IsButtonAvailable(ID) = True Then
            Dim Ptr As LongPtr
            Ptr = GetButtonPtr(ID)
            If Ptr <> NULL_PTR Then Set InsertMark = PtrToObj(Ptr)
        End If
        After = CBool(.dwFlags = TBIMHT_AFTER)
    End If
    End With
End If
End Property

Public Property Let InsertMark(Optional ByRef After As Boolean, ByVal Value As TbrButton)
Set Me.InsertMark(After) = Value
End Property

Public Property Set InsertMark(Optional ByRef After As Boolean, ByVal Value As TbrButton)
If ToolBarHandle <> NULL_PTR Then
    Dim TBIM As TBINSERTMARK
    With TBIM
    If Value Is Nothing Then
        .iButton = -1
        .dwFlags = 0
    Else
        .iButton = Value.Position - 1
        .dwFlags = IIf(After = True, TBIMHT_AFTER, 0)
    End If
    End With
    SendMessage ToolBarHandle, TB_SETINSERTMARK, 0, ByVal VarPtr(TBIM)
End If
End Property

Public Property Get HotItem() As TbrButton
Attribute HotItem.VB_Description = "Returns/sets a reference to the currently hot button."
If ToolBarHandle <> NULL_PTR Then
    Dim Index As Long
    Index = CLng(SendMessage(ToolBarHandle, TB_GETHOTITEM, 0, ByVal 0&) + 1)
    If Index > 0 Then
        Dim Ptr As LongPtr
        Ptr = GetButtonPtr(GetButtonID(Index))
        If Ptr <> NULL_PTR Then Set HotItem = PtrToObj(Ptr)
    End If
End If
End Property

Public Property Let HotItem(ByVal Value As TbrButton)
Set Me.HotItem = Value
End Property

Public Property Set HotItem(ByVal Value As TbrButton)
If ToolBarHandle <> NULL_PTR Then
    If Not Value Is Nothing Then
        Value.Hot = True
    Else
        SendMessage ToolBarHandle, TB_SETHOTITEM, -1, ByVal 0&
    End If
End If
End Property

Public Property Get ImageHeight() As Single
Attribute ImageHeight.VB_Description = "Returns/sets the height of the images."
Attribute ImageHeight.VB_MemberFlags = "400"
ImageHeight = UserControl.ScaleY(HiWord(ToolBarImageSize), vbPixels, vbContainerSize)
End Property

Public Property Let ImageHeight(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim IntValue As Integer
On Error Resume Next
IntValue = CInt(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
If Err.Number <> 0 Then IntValue = 0
On Error GoTo 0
ToolBarImageSize = MakeDWord(LoWord(ToolBarImageSize), IntValue)
If ToolBarHandle <> NULL_PTR Then
    SendMessage ToolBarHandle, TB_SETBITMAPSIZE, 0, ByVal ToolBarImageSize
    Call ReCreateButtons
End If
Call UserControl_Resize
End Property

Public Property Get ImageWidth() As Single
Attribute ImageWidth.VB_Description = "Returns/sets the width of the images."
Attribute ImageWidth.VB_MemberFlags = "400"
ImageWidth = UserControl.ScaleY(LoWord(ToolBarImageSize), vbPixels, vbContainerSize)
End Property

Public Property Let ImageWidth(ByVal Value As Single)
If Value < 0 Then Err.Raise 380
Dim IntValue As Integer
On Error Resume Next
IntValue = CInt(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
If Err.Number <> 0 Then IntValue = 0
On Error GoTo 0
ToolBarImageSize = MakeDWord(IntValue, HiWord(ToolBarImageSize))
If ToolBarHandle <> NULL_PTR Then
    SendMessage ToolBarHandle, TB_SETBITMAPSIZE, 0, ByVal ToolBarImageSize
    Call ReCreateButtons
End If
Call UserControl_Resize
End Property

Private Sub ResetCustomizeButtons()
If ToolBarCustomizeButtonsCount = 0 Then Exit Sub
If ToolBarHandle <> NULL_PTR Then
    Do While SendMessage(ToolBarHandle, TB_DELETEBUTTON, 0, ByVal 0&) <> 0: Loop
    Dim TBB() As TBBUTTON, i As Long
    ReDim TBB(0 To (ToolBarCustomizeButtonsCount - 1)) As TBBUTTON
    For i = 1 To ToolBarCustomizeButtonsCount
        LSet TBB(i - 1) = ToolBarCustomizeButtons(i).TBB
        TBB(i - 1).iString = StrPtr(ToolBarCustomizeButtons(i).Caption)
    Next i
    SendMessage ToolBarHandle, TB_ADDBUTTONS, ToolBarCustomizeButtonsCount, ByVal VarPtr(TBB(0))
    Dim TBBI As TBBUTTONINFO
    TBBI.cbSize = LenB(TBBI)
    TBBI.dwMask = TBIF_SIZE
    For i = 1 To ToolBarCustomizeButtonsCount
        With ToolBarCustomizeButtons(i)
        If .CX > 0 And (.TBB.fsStyle And BTNS_AUTOSIZE) = 0 Then
            TBBI.CX = .CX
            SendMessage ToolBarHandle, TB_SETBUTTONINFO, GetButtonID(i), ByVal VarPtr(TBBI)
        End If
        End With
    Next i
    ToolBarCustomizeButtonsCount = 0
End If
End Sub

Private Sub AllocCustomizeButtons()
If ToolBarCustomizeButtonsCount > 0 Then Exit Sub
If ToolBarHandle <> NULL_PTR Then
    Erase ToolBarCustomizeButtons()
    ToolBarCustomizeButtonsCount = CLng(SendMessage(ToolBarHandle, TB_BUTTONCOUNT, 0, ByVal 0&))
    If ToolBarCustomizeButtonsCount > 0 Then
        ReDim ToolBarCustomizeButtons(1 To ToolBarCustomizeButtonsCount) As ShadowButtonStruct
        Dim i As Long
        For i = 1 To ToolBarCustomizeButtonsCount: ToolBarCustomizeButtons(i) = GetShadowButton(GetButtonID(i)): Next i
    End If
End If
End Sub

Private Sub ModifyButton(ByVal ID As Long, ByRef NewButton As ShadowButtonStruct)
If ToolBarHandle <> NULL_PTR Then
    Dim Index As Long
    Index = CLng(SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&)) + 1
    If Index > 0 Then
        Dim Count As Long, i As Long, NoText As Boolean
        Count = CLng(SendMessage(ToolBarHandle, TB_BUTTONCOUNT, 0, ByVal 0&))
        NoText = True
        For i = 1 To Count
            If Not GetButtonText(GetButtonID(i)) = vbNullString Then NoText = False
        Next i
        If NoText = False Then SendMessage ToolBarHandle, TB_SETBUTTONSIZE, 0, ByVal MakeDWord(PropButtonWidth + 10, PropButtonHeight + 10)
        SendMessage ToolBarHandle, TB_DELETEBUTTON, Index - 1, ByVal 0&
        With NewButton
        Dim TBB As TBBUTTON
        LSet TBB = .TBB
        TBB.iString = StrPtr(.Caption)
        SendMessage ToolBarHandle, TB_INSERTBUTTON, Index - 1, ByVal VarPtr(TBB)
        If .CX > 0 And (.TBB.fsStyle And BTNS_AUTOSIZE) = 0 Then
            Dim TBBI As TBBUTTONINFO
            TBBI.cbSize = LenB(TBBI)
            TBBI.dwMask = TBIF_SIZE
            TBBI.CX = .CX
            SendMessage ToolBarHandle, TB_SETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        End If
        End With
        Dim Size As Long
        Size = CLng(SendMessage(ToolBarHandle, TB_GETBUTTONSIZE, 0, ByVal 0&))
        PropButtonWidth = LoWord(Size)
        PropButtonHeight = HiWord(Size)
    End If
End If
End Sub

Private Sub ReCreateButtons()
If ToolBarHandle <> NULL_PTR Then
    Dim Count As Long
    Count = CLng(SendMessage(ToolBarHandle, TB_BUTTONCOUNT, 0, ByVal 0&))
    If Count > 0 Then
        Dim ReButtons() As ShadowButtonStruct
        ReDim ReButtons(1 To Count) As ShadowButtonStruct
        Dim i As Long, NoText As Boolean
        NoText = True
        For i = 1 To Count
            ReButtons(i) = GetShadowButton(GetButtonID(i))
            If Not ReButtons(i).Caption = vbNullString Then NoText = False
        Next i
        If NoText = False Then SendMessage ToolBarHandle, TB_SETBUTTONSIZE, 0, ByVal MakeDWord(PropButtonWidth + 10, PropButtonHeight + 10)
        Do While SendMessage(ToolBarHandle, TB_DELETEBUTTON, 0, ByVal 0&) <> 0: Loop
        Dim TBB() As TBBUTTON
        ReDim TBB(0 To (Count - 1)) As TBBUTTON
        For i = 0 To (Count - 1)
            With ReButtons(i + 1)
            LSet TBB(i) = .TBB
            TBB(i).iString = StrPtr(.Caption)
            End With
        Next i
        SendMessage ToolBarHandle, TB_ADDBUTTONS, Count, ByVal VarPtr(TBB(0))
        Dim TBBI As TBBUTTONINFO
        TBBI.cbSize = LenB(TBBI)
        TBBI.dwMask = TBIF_SIZE
        For i = 1 To Count
            With ReButtons(i)
            If .CX > 0 And (.TBB.fsStyle And BTNS_AUTOSIZE) = 0 Then
                TBBI.CX = .CX
                SendMessage ToolBarHandle, TB_SETBUTTONINFO, .TBB.IDCommand, ByVal VarPtr(TBBI)
            End If
            End With
        Next i
        Dim Size As Long
        Size = CLng(SendMessage(ToolBarHandle, TB_GETBUTTONSIZE, 0, ByVal 0&))
        PropButtonWidth = LoWord(Size)
        PropButtonHeight = HiWord(Size)
    End If
    InvalidateRect ToolBarHandle, ByVal NULL_PTR, 1
End If
End Sub

Private Function GetButtonID(ByVal Index As Long) As Long
If ToolBarHandle <> NULL_PTR Then
    Dim TBB As TBBUTTON
    SendMessage ToolBarHandle, TB_GETBUTTON, Index - 1, ByVal VarPtr(TBB)
    GetButtonID = TBB.IDCommand
End If
End Function

Private Function GetButtonPtr(ByVal ID As Long) As LongPtr
If ToolBarHandle <> NULL_PTR And IsButtonAvailable(ID) = True Then
    Dim TBBI As TBBUTTONINFO
    With TBBI
    .cbSize = LenB(TBBI)
    .dwMask = TBIF_LPARAM
    SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
    GetButtonPtr = .lParam
    End With
End If
End Function

Private Function GetButtonText(ByVal ID As Long) As String
If ToolBarHandle <> NULL_PTR Then
    Dim Length As Long
    Length = CLng(SendMessage(ToolBarHandle, TB_GETBUTTONTEXT, ID, ByVal 0&))
    If Length > 0 Then
        Dim Buffer As String
        Buffer = String(Length, vbNullChar) & vbNullChar
        SendMessage ToolBarHandle, TB_GETBUTTONTEXT, ID, ByVal StrPtr(Buffer)
        GetButtonText = Left$(Buffer, Length)
    End If
End If
End Function

Private Function IsButtonAvailable(ByVal ID As Long) As Boolean
If ToolBarHandle <> NULL_PTR Then IsButtonAvailable = CBool(SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&) > -1)
End Function

Private Function GetShadowButton(ByVal ID As Long) As ShadowButtonStruct
If ToolBarHandle <> NULL_PTR Then
    Dim Index As Long
    Index = CLng(SendMessage(ToolBarHandle, TB_COMMANDTOINDEX, ID, ByVal 0&)) + 1
    If Index > 0 Then
        With GetShadowButton
        SendMessage ToolBarHandle, TB_GETBUTTON, Index - 1, ByVal VarPtr(.TBB)
        .Caption = GetButtonText(ID)
        Dim TBBI As TBBUTTONINFO
        TBBI.cbSize = LenB(TBBI)
        TBBI.dwMask = TBIF_SIZE
        SendMessage ToolBarHandle, TB_GETBUTTONINFO, ID, ByVal VarPtr(TBBI)
        .CX = TBBI.CX
        End With
    End If
End If
End Function

Private Function ShowButtonMenuItems(ByVal Button As TbrButton, ByVal Keyboard As Boolean) As Long
If ToolBarHandle <> NULL_PTR Then
    If ToolBarPopupMenuHandle <> NULL_PTR Then
        SendMessage ToolBarHandle, WM_CANCELMODE, 0, ByVal 0&
        Exit Function
    End If
    If Button.ButtonMenus.Count > 0 Then
        Dim Text As String, Count As Long, MenuItem As Long, HasMenuPictureCallback As Boolean
        ToolBarPopupMenuHandle = CreatePopupMenu()
        Dim TPMP As TPMPARAMS, P As POINTAPI, MII As MENUITEMINFO
        TPMP.cbSize = LenB(TPMP)
        SendMessage ToolBarHandle, TB_GETRECT, Button.ID, ByVal VarPtr(TPMP.RCExclude)
        MapWindowPoints ToolBarHandle, HWND_DESKTOP, TPMP.RCExclude, 2
        P.X = TPMP.RCExclude.Left
        P.Y = TPMP.RCExclude.Bottom
        MII.cbSize = LenB(MII)
        For MenuItem = 1 To Button.ButtonMenus.Count
            With Button.ButtonMenus(MenuItem)
            If .Visible = True Then
                If .Separator = False Then
                    MII.fMask = MIIM_STATE Or MIIM_ID Or MIIM_STRING
                    MII.fType = 0
                    Text = .Text
                    MII.dwTypeData = StrPtr(Text)
                    MII.cch = Len(Text)
                    If .Picture Is Nothing Then
                        MII.hBmpItem = NULL_PTR
                    ElseIf .Picture.Handle = NULL_PTR Then
                        MII.hBmpItem = NULL_PTR
                    Else
                        ' The menu theme is removed when some menu item has hBmpItem set to HBMMENU_CALLBACK.
                        ' Use 32-bit pre-multiplied alpha RGB bitmaps for best results.
                        MII.fMask = MII.fMask Or MIIM_BITMAP
                        If .Picture.Type = vbPicTypeBitmap Then
                            MII.hBmpItem = .Picture.Handle
                        Else
                            MII.hBmpItem = HBMMENU_CALLBACK
                            HasMenuPictureCallback = True
                        End If
                    End If
                    If .Enabled = True Then
                        MII.fState = MFS_ENABLED
                    Else
                        MII.fState = MFS_DISABLED
                    End If
                    If .Checked = True Then
                        MII.fState = MII.fState Or MFS_CHECKED
                    Else
                        MII.fState = MII.fState Or MFS_UNCHECKED
                    End If
                Else
                    MII.fMask = MIIM_STATE Or MIIM_ID Or MIIM_FTYPE
                    MII.fType = MFT_SEPARATOR
                    MII.dwTypeData = 0
                    MII.cch = 0
                    MII.hBmpItem = NULL_PTR
                End If
                MII.wID = MenuItem
                InsertMenuItem ToolBarPopupMenuHandle, 0, 0, MII
                Count = Count + 1
            End If
            End With
        Next MenuItem
        If Count > 0 Then
            Dim MI As MENUINFO
            MI.cbSize = LenB(MI)
            MI.fMask = MIM_MENUDATA
            MI.dwMenuData = ObjPtr(Button)
            If HasMenuPictureCallback = True Then
                ' The menu theme is lost due to HBMMENU_CALLBACK.
                ' Setting a menu background color fixes a one-pixel overlap between the picture and text.
                MI.fMask = MI.fMask Or MIM_BACKGROUND
                MI.hBrBack = GetSysColorBrush(COLOR_MENU)
            End If
            SetMenuInfo ToolBarPopupMenuHandle, MI
            Dim Flags As Long
            If PropRightToLeft = False Then
                Flags = TPM_LEFTALIGN
            Else
                If PropRightToLeftLayout = True Then Flags = TPM_RIGHTALIGN Else Flags = TPM_LEFTALIGN Or TPM_LAYOUTRTL
            End If
            Flags = Flags Or TPM_TOPALIGN Or TPM_LEFTBUTTON Or TPM_VERTICAL Or TPM_RETURNCMD
            Set ToolBarPopupMenuButton = Button
            ToolBarPopupMenuKeyboard = Keyboard
            ShowButtonMenuItems = TrackPopupMenuEx(ToolBarPopupMenuHandle, Flags, P.X, P.Y, ToolBarHandle, TPMP)
        End If
        DestroyMenu ToolBarPopupMenuHandle
        ToolBarPopupMenuHandle = NULL_PTR
        Set ToolBarPopupMenuButton = Nothing
        ToolBarPopupMenuKeyboard = False
    End If
End If
End Function

Private Function NextButtonID() As Long
Static ID As Long
ID = ID + 1
NextButtonID = ID
End Function

Private Function ImageListSizesAreEqual() As Boolean
ImageListSizesAreEqual = True
Dim CheckSize As SIZEAPI, i As Long
For i = 1 To 4
    CheckSize.CX = VBA.Choose(i, ImageListSize.CX, DisabledImageListSize.CX, HotImageListSize.CX, PressedImageListSize.CX)
    CheckSize.CY = VBA.Choose(i, ImageListSize.CY, DisabledImageListSize.CY, HotImageListSize.CY, PressedImageListSize.CY)
    If CheckSize.CX > 0 And CheckSize.CY > 0 Then
        If ImageListSize.CX > 0 And ImageListSize.CY > 0 Then
            If CheckSize.CX <> ImageListSize.CX Or CheckSize.CY <> ImageListSize.CY Then ImageListSizesAreEqual = False
        End If
        If DisabledImageListSize.CX > 0 And DisabledImageListSize.CY > 0 Then
            If CheckSize.CX <> DisabledImageListSize.CX Or CheckSize.CY <> DisabledImageListSize.CY Then ImageListSizesAreEqual = False
        End If
        If HotImageListSize.CX > 0 And HotImageListSize.CY > 0 Then
            If CheckSize.CX <> HotImageListSize.CX Or CheckSize.CY <> HotImageListSize.CY Then ImageListSizesAreEqual = False
        End If
        If PressedImageListSize.CX > 0 And PressedImageListSize.CY > 0 Then
            If CheckSize.CX <> PressedImageListSize.CX Or CheckSize.CY <> PressedImageListSize.CY Then ImageListSizesAreEqual = False
        End If
    End If
Next i
End Function

Private Function CreateTransparentBrush(ByVal hDC As LongPtr) As LongPtr
Dim hDCBmp As LongPtr
Dim hBmp As LongPtr, hBmpOld As LongPtr
With UserControl
hDCBmp = CreateCompatibleDC(hDC)
If hDCBmp <> NULL_PTR Then
    hBmp = CreateCompatibleBitmap(hDC, .ScaleWidth, .ScaleHeight)
    If hBmp <> NULL_PTR Then
        hBmpOld = SelectObject(hDCBmp, hBmp)
        Dim WndRect As RECT, P As POINTAPI
        GetWindowRect .hWnd, WndRect
        MapWindowPoints HWND_DESKTOP, .ContainerHwnd, WndRect, 2
        P.X = WndRect.Left
        P.Y = WndRect.Top
        SetViewportOrgEx hDCBmp, -P.X, -P.Y, P
        SendMessage .ContainerHwnd, WM_PAINT, hDCBmp, ByVal 0&
        SetViewportOrgEx hDCBmp, P.X, P.Y, P
        CreateTransparentBrush = CreatePatternBrush(hBmp)
        SelectObject hDCBmp, hBmpOld
        DeleteObject hBmp
    End If
    DeleteDC hDCBmp
End If
End With
End Function

Private Sub SetVisualStylesToolTip()
If ToolBarHandle <> NULL_PTR Then
    If ToolBarToolTipHandle <> NULL_PTR And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles ToolBarToolTipHandle
        Else
            RemoveVisualStyles ToolBarToolTipHandle
        End If
    End If
End If
End Sub

Private Function PropImageListControl() As Object
If ToolBarImageListObjectPointer <> NULL_PTR Then Set PropImageListControl = PtrToObj(ToolBarImageListObjectPointer)
End Function

Private Function PropDisabledImageListControl() As Object
If ToolBarDisabledImageListObjectPointer <> NULL_PTR Then Set PropDisabledImageListControl = PtrToObj(ToolBarDisabledImageListObjectPointer)
End Function

Private Function PropHotImageListControl() As Object
If ToolBarHotImageListObjectPointer <> NULL_PTR Then Set PropHotImageListControl = PtrToObj(ToolBarHotImageListObjectPointer)
End Function

Private Function PropPressedImageListControl() As Object
If ToolBarPressedImageListObjectPointer <> NULL_PTR Then Set PropPressedImageListControl = PtrToObj(ToolBarPressedImageListObjectPointer)
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
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
    Case 3
        ISubclass_Message = WindowProcUserControlDesignMode(hWnd, wMsg, wParam, lParam)
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
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
    Case WM_ERASEBKGND
        If PropDoubleBuffer = True And (ToolBarDoubleBufferEraseBkgDC <> wParam Or ToolBarDoubleBufferEraseBkgDC = NULL_PTR) And WindowFromDC(wParam) = hWnd Then
            WindowProcControl = 0
            Exit Function
        Else
            If ToolBarHandle <> NULL_PTR Then
                If ToolBarBackColorBrush <> NULL_PTR And (SendMessage(ToolBarHandle, TB_GETEXTENDEDSTYLE, 0, ByVal 0&) And TBSTYLE_EX_DOUBLEBUFFER) = 0 Then
                    Dim RC As RECT
                    GetClientRect hWnd, RC
                    FillRect wParam, RC, ToolBarBackColorBrush
                    If PropTransparent = True Then
                        If ToolBarTransparentBrush = NULL_PTR Then ToolBarTransparentBrush = CreateTransparentBrush(wParam)
                        If ToolBarTransparentBrush <> NULL_PTR Then FillRect wParam, RC, ToolBarTransparentBrush
                    End If
                    WindowProcControl = 1
                    Exit Function
                End If
            End If
        End If
    Case WM_PAINT
        If PropDoubleBuffer = True Then
            Dim ClientRect As RECT, hDC As LongPtr
            Dim hDCBmp As LongPtr
            Dim hBmp As LongPtr, hBmpOld As LongPtr
            GetClientRect hWnd, ClientRect
            Dim PS As PAINTSTRUCT
            hDC = BeginPaint(hWnd, PS)
            With PS
            If wParam <> 0 Then hDC = wParam
            hDCBmp = CreateCompatibleDC(hDC)
            If hDCBmp <> NULL_PTR Then
                hBmp = CreateCompatibleBitmap(hDC, ClientRect.Right - ClientRect.Left, ClientRect.Bottom - ClientRect.Top)
                If hBmp <> NULL_PTR Then
                    hBmpOld = SelectObject(hDCBmp, hBmp)
                    ToolBarDoubleBufferEraseBkgDC = hDCBmp
                    SendMessage hWnd, WM_PRINT, hDCBmp, ByVal PRF_CLIENT Or PRF_ERASEBKGND
                    ToolBarDoubleBufferEraseBkgDC = NULL_PTR
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
    Case WM_ENTERMENULOOP
        If ToolBarPopupMenuHandle <> NULL_PTR And wParam = 1 And ToolBarPopupMenuKeyboard = True Then
            Const INPUT_KEYBOARD As Long = 1, KEYEVENTF_KEYUP As Long = &H2
            Dim GI As GENERALINPUT
            GI.dwType = INPUT_KEYBOARD
            With GI.KEYBDI
            .wVKey = vbKeyDown
            SendInput 1, GI, LenB(GI)
            .dwFlags = KEYEVENTF_KEYUP
            SendInput 1, GI, LenB(GI)
            End With
        End If
    Case WM_MEASUREITEM, WM_DRAWITEM
        If ToolBarPopupMenuHandle <> NULL_PTR And Not ToolBarPopupMenuButton Is Nothing Then
            Dim MenuPicture As IPictureDisp, CX As Long, CY As Long
            Select Case wMsg
                Case WM_MEASUREITEM
                    Dim MIS As MEASUREITEMSTRUCT
                    CopyMemory MIS, ByVal lParam, LenB(MIS)
                    If MIS.CtlType = ODT_MENU And MIS.ItemID >= 1 And MIS.ItemID <= ToolBarPopupMenuButton.ButtonMenus.Count Then
                        Set MenuPicture = ToolBarPopupMenuButton.ButtonMenus(MIS.ItemID).Picture
                        If Not MenuPicture Is Nothing Then
                            CX = CHimetricToPixel_X(MenuPicture.Width)
                            CY = CHimetricToPixel_Y(MenuPicture.Height)
                            MIS.ItemWidth = MIS.ItemWidth + CX
                            If MIS.ItemHeight < CY Then MIS.ItemHeight = CY
                            CopyMemory ByVal lParam, MIS, LenB(MIS)
                            WindowProcControl = 1
                            Exit Function
                        End If
                    End If
                Case WM_DRAWITEM
                    Dim DIS As DRAWITEMSTRUCT
                    CopyMemory DIS, ByVal lParam, LenB(DIS)
                    If DIS.CtlType = ODT_MENU And DIS.hWndItem = ToolBarPopupMenuHandle And DIS.ItemID >= 1 And DIS.ItemID <= ToolBarPopupMenuButton.ButtonMenus.Count Then
                        Set MenuPicture = ToolBarPopupMenuButton.ButtonMenus(DIS.ItemID).Picture
                        If Not MenuPicture Is Nothing Then
                            CX = CHimetricToPixel_X(MenuPicture.Width)
                            CY = CHimetricToPixel_Y(MenuPicture.Height)
                            If Not (DIS.ItemState And ODS_DISABLED) = ODS_DISABLED Then
                                Call RenderPicture(MenuPicture, DIS.hDC, DIS.RCItem.Left, DIS.RCItem.Top + ((DIS.RCItem.Bottom - DIS.RCItem.Top - CY) / 2), CX, CY, 1)
                            Else
                                If MenuPicture.Type = vbPicTypeIcon Then
                                    DrawState DIS.hDC, NULL_PTR, NULL_PTR, MenuPicture.Handle, 0, DIS.RCItem.Left, DIS.RCItem.Top + ((DIS.RCItem.Bottom - DIS.RCItem.Top - CY) / 2), CX, CY, DST_ICON Or DSS_DISABLED
                                Else
                                    Dim hImage As LongPtr
                                    hImage = BitmapHandleFromPicture(MenuPicture, vbWhite)
                                    ' The DrawState API with DSS_DISABLED will draw white as transparent.
                                    ' This will ensure GIF bitmaps or metafiles are better drawn.
                                    DrawState DIS.hDC, NULL_PTR, NULL_PTR, hImage, 0, DIS.RCItem.Left, DIS.RCItem.Top + ((DIS.RCItem.Bottom - DIS.RCItem.Top - CY) / 2), CX, CY, DST_BITMAP Or DSS_DISABLED
                                    DeleteObject hImage
                                End If
                            End If
                            WindowProcControl = 1
                            Exit Function
                        End If
                    End If
            End Select
        End If
    Case UM_SETBUTTONCX
        If wParam > 0 And lParam > 0 Then
            Dim TBBI As TBBUTTONINFO
            With TBBI
            .cbSize = LenB(TBBI)
            .dwMask = TBIF_SIZE
            .CX = CLng(lParam)
            End With
            SendMessage ToolBarHandle, TB_SETBUTTONINFO, wParam, ByVal VarPtr(TBBI)
        End If
    Case WM_UPDATEUISTATE
        ' When a ToolBar is hosted in a MDIForm it *can* happen that an MDIChild sets UISF_HIDEACCEL to it's owner.
        ' However, this ensures to circumvent such scenario.
        If LoWord(CLng(wParam)) = UIS_SET Then
            Dim IntValue As Integer
            IntValue = HiWord(CLng(wParam))
            If (IntValue And UISF_HIDEACCEL) = UISF_HIDEACCEL Then
                IntValue = IntValue And Not UISF_HIDEACCEL
                wParam = MakeDWord(UIS_SET, IntValue)
            End If
        End If
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
        ' Necessary to process here as NM_DBLCLK will not be fired. (Bug?)
        ' Though NM_RDBLCLK will be fired.
        RaiseEvent DblClick
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
        Dim X As Single
        Dim Y As Single
        X = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
        Y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
        Select Case wMsg
            Case WM_LBUTTONDOWN
                RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
                ToolBarIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                ToolBarIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                ToolBarIsClick = True
            Case WM_MOUSEMOVE
                Dim P As POINTAPI, Index As Long
                If PropHotTracking = True And PropStyle = TbrStyleStandard Then
                    P.X = Get_X_lParam(lParam)
                    P.Y = Get_Y_lParam(lParam)
                    Index = CLng(SendMessage(ToolBarHandle, TB_HITTEST, 0, ByVal VarPtr(P))) + 1
                    If SendMessage(ToolBarHandle, TB_GETANCHORHIGHLIGHT, 0, ByVal 0&) = 0 Then
                        SendMessage ToolBarHandle, TB_SETHOTITEM, Index - 1, ByVal 0&
                    Else
                        If Index > 0 Then SendMessage ToolBarHandle, TB_SETHOTITEM, Index - 1, ByVal 0&
                    End If
                End If
                If ToolBarMouseOver = False And PropMouseTrack = True Then
                    ToolBarMouseOver = True
                    RaiseEvent MouseEnter
                    P.X = Get_X_lParam(lParam)
                    P.Y = Get_Y_lParam(lParam)
                    ToolBarMouseOverIndex = CLng(SendMessage(ToolBarHandle, TB_HITTEST, 0, ByVal VarPtr(P))) + 1
                    If ToolBarMouseOverIndex > 0 Then
                        Dim ID1 As Long
                        ID1 = GetButtonID(ToolBarMouseOverIndex)
                        If IsButtonAvailable(ID1) = True Then
                            Dim Ptr1 As LongPtr
                            Ptr1 = GetButtonPtr(ID1)
                            If Ptr1 <> NULL_PTR Then RaiseEvent ButtonMouseEnter(PtrToObj(Ptr1))
                        End If
                    End If
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                If ToolBarMouseOver = True And PropMouseTrack = True Then
                    P.X = Get_X_lParam(lParam)
                    P.Y = Get_Y_lParam(lParam)
                    Index = CLng(SendMessage(ToolBarHandle, TB_HITTEST, 0, ByVal VarPtr(P))) + 1
                    If ToolBarMouseOverIndex <> Index Then
                        If ToolBarMouseOverIndex > 0 Then
                            Dim ID2 As Long
                            ID2 = GetButtonID(ToolBarMouseOverIndex)
                            If IsButtonAvailable(ID2) = True Then
                                Dim Ptr2 As LongPtr
                                Ptr2 = GetButtonPtr(ID2)
                                If Ptr2 <> NULL_PTR Then RaiseEvent ButtonMouseLeave(PtrToObj(Ptr2))
                            End If
                        End If
                        ToolBarMouseOverIndex = Index
                        If ToolBarMouseOverIndex > 0 Then
                            Dim ID3 As Long
                            ID3 = GetButtonID(ToolBarMouseOverIndex)
                            If IsButtonAvailable(ID3) = True Then
                                Dim Ptr3 As LongPtr
                                Ptr3 = GetButtonPtr(ID3)
                                If Ptr3 <> NULL_PTR Then RaiseEvent ButtonMouseEnter(PtrToObj(Ptr3))
                            End If
                        End If
                    End If
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
                If ToolBarIsClick = True Then
                    ToolBarIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If ToolBarMouseOver = True Then
            ToolBarMouseOver = False
            If ToolBarMouseOverIndex > 0 Then
                Dim ID4 As Long
                ID4 = GetButtonID(ToolBarMouseOverIndex)
                If IsButtonAvailable(ID4) = True Then
                    Dim Ptr4 As LongPtr
                    Ptr4 = GetButtonPtr(ID4)
                    If Ptr4 <> NULL_PTR Then RaiseEvent ButtonMouseLeave(PtrToObj(Ptr4))
                End If
            End If
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_COMMAND
        Const BN_CLICKED As Long = 0
        If lParam = ToolBarHandle And HiWord(CLng(wParam)) = BN_CLICKED Then
            ' This notification is normally ignored by whole drop down buttons.
            ' Anyhow, it will be sent when pressing RBUTTONDOWN-LBUTTONDOWN-LBUTTONUP. (Bug?)
            ' Therefore it is necessary to check whether the style BTNS_WHOLEDROPDOWN is set or not.
            Dim TBBI As TBBUTTONINFO
            With TBBI
            .cbSize = LenB(TBBI)
            .dwMask = TBIF_LPARAM Or TBIF_STYLE
            SendMessage ToolBarHandle, TB_GETBUTTONINFO, LoWord(CLng(wParam)), ByVal VarPtr(TBBI)
            If .lParam <> 0 And (.fsStyle And BTNS_WHOLEDROPDOWN) = 0 Then RaiseEvent ButtonClick(PtrToObj(.lParam))
            End With
        End If
    Case WM_ERASEBKGND
        If ToolBarHandle <> NULL_PTR Then
            If ToolBarBackColorBrush <> NULL_PTR And (SendMessage(ToolBarHandle, TB_GETEXTENDEDSTYLE, 0, ByVal 0&) And TBSTYLE_EX_DOUBLEBUFFER) = 0 Then
                Dim RC As RECT
                GetClientRect hWnd, RC
                FillRect wParam, RC, ToolBarBackColorBrush
                If PropTransparent = True Then
                    If ToolBarTransparentBrush = NULL_PTR Then ToolBarTransparentBrush = CreateTransparentBrush(wParam)
                    If ToolBarTransparentBrush <> NULL_PTR Then FillRect wParam, RC, ToolBarTransparentBrush
                End If
                WindowProcUserControl = 1
            End If
            ' It is necessary to exit this message in all cases to avoid artifacts. (Bug?)
            Exit Function
        End If
    Case WM_NOTIFY
        Dim NM As NMHDR, NMTB As NMTOOLBAR, Button As TbrButton
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = ToolBarHandle Then
            Select Case NM.Code
                Case NM_CUSTOMDRAW
                    Dim NMTBCD As NMTBCUSTOMDRAW
                    CopyMemory NMTBCD, ByVal lParam, LenB(NMTBCD)
                    Select Case NMTBCD.NMCD.dwDrawStage
                        Case CDDS_PREPAINT
                            WindowProcUserControl = CDRF_NOTIFYITEMDRAW
                            Exit Function
                        Case CDDS_ITEMPREPAINT
                            WindowProcUserControl = CDRF_DODEFAULT
                            If NMTBCD.NMCD.lItemlParam <> 0 Then
                                Set Button = PtrToObj(NMTBCD.NMCD.lItemlParam)
                                If (NMTBCD.NMCD.uItemState And CDIS_MARKED) = CDIS_MARKED Then WindowProcUserControl = WindowProcUserControl Or TBCDRF_BLENDICON
                                If PropHotTracking = False Or (NMTBCD.NMCD.uItemState And CDIS_HOT) = 0 Then
                                    NMTBCD.ClrText = WinColor(Button.ForeColor)
                                Else
                                    NMTBCD.ClrText = WinColor(vbHighlightText)
                                    NMTBCD.ClrHighlightHotTrack = WinColor(vbHighlight)
                                    If PropStyle = TbrStyleStandard Then NMTBCD.ClrBtnFace = WinColor(vbHighlight)
                                    WindowProcUserControl = WindowProcUserControl Or TBCDRF_HILITEHOTTRACK
                                End If
                                CopyMemory ByVal lParam, NMTBCD, LenB(NMTBCD)
                                If ComCtlsSupportLevel() >= 2 Then WindowProcUserControl = WindowProcUserControl Or TBCDRF_USECDCOLORS
                            End If
                            Exit Function
                    End Select
                Case TBN_GETINFOTIP
                    Dim NMTBGIT As NMTBGETINFOTIP
                    CopyMemory NMTBGIT, ByVal lParam, LenB(NMTBGIT)
                    With NMTBGIT
                    If .iItem > 0 And .lParam <> 0 And .pszText <> NULL_PTR Then
                        Set Button = PtrToObj(.lParam)
                        Dim ToolTipText As String
                        ToolTipText = Button.ToolTipText
                        If Not ToolTipText = vbNullString Then
                            If PropRightToLeft = True And PropRightToLeftLayout = False Then ToolTipText = ChrW(&H202B) & ToolTipText ' Right-to-left Embedding (RLE)
                            ToolTipText = Left$(ToolTipText, .cchTextMax - 1) & vbNullChar
                            CopyMemory ByVal .pszText, ByVal StrPtr(ToolTipText), LenB(ToolTipText)
                        Else
                            CopyMemory ByVal .pszText, 0&, 4
                        End If
                    End If
                    End With
                Case TBN_DRAGOUT
                    CopyMemory NMTB, ByVal lParam, LenB(NMTB)
                    If NMTB.iItem > 0 Then
                        Set Button = PtrToObj(GetButtonPtr(NMTB.iItem))
                        RaiseEvent ButtonDrag(Button, GetMouseStateFromMsg())
                    End If
                Case TBN_DROPDOWN
                    CopyMemory NMTB, ByVal lParam, LenB(NMTB)
                    If NMTB.iItem > 0 Then
                        Set Button = PtrToObj(GetButtonPtr(NMTB.iItem))
                        RaiseEvent ButtonDropDown(Button)
                        Dim MenuItem As Long
                        MenuItem = ShowButtonMenuItems(Button, False)
                        If MenuItem >= 1 And MenuItem <= Button.ButtonMenus.Count Then RaiseEvent ButtonMenuClick(Button.ButtonMenus(MenuItem))
                        WindowProcUserControl = TBDDRET_DEFAULT
                    Else
                        WindowProcUserControl = TBDDRET_NODEFAULT
                    End If
                    Exit Function
                Case TBN_HOTITEMCHANGE
                    Dim NMTBHI As NMTBHOTITEM
                    CopyMemory NMTBHI, ByVal lParam, LenB(NMTBHI)
                    If NMTBHI.IDNew > 0 Then RaiseEvent ButtonHotChanged(PtrToObj(GetButtonPtr(NMTBHI.IDNew)), True)
                    If NMTBHI.IDOld > 0 Then RaiseEvent ButtonHotChanged(PtrToObj(GetButtonPtr(NMTBHI.IDOld)), False)
                Case TBN_BEGINADJUST
                    Call AllocCustomizeButtons
                    RaiseEvent BeginCustomization
                Case TBN_INITCUSTOMIZE
                    Dim hDlg32 As Long, HideHelpButton As Boolean
                    ' lParam points to a struct of this kind: (undocumented)
                    ' hdr As NMHDR
                    ' hDlg As LongPtr
                    CopyMemory hDlg32, ByVal UnsignedAdd(lParam, LenB(NM)), 4
                    RaiseEvent InitCustomizationDialog(hDlg32, HideHelpButton)
                    If HideHelpButton = True Then
                        WindowProcUserControl = TBNRF_HIDEHELP
                    Else
                        WindowProcUserControl = 0
                    End If
                    Exit Function
                Case TBN_QUERYINSERT, TBN_QUERYDELETE
                    WindowProcUserControl = 1
                    Exit Function
                Case TBN_SAVE
                    Dim NMTBS As NMTBSAVE
                    CopyMemory NMTBS, ByVal lParam, LenB(NMTBS)
                    If NMTBS.iItem = -1 Then Call AllocCustomizeButtons
                Case TBN_RESTORE
                    Dim NMTBR As NMTBRESTORE
                    CopyMemory NMTBR, ByVal lParam, LenB(NMTBR)
                    If NMTBR.iItem = -1 Then Call AllocCustomizeButtons
                Case TBN_TOOLBARCHANGE
                    RaiseEvent CustomizationChange
                Case TBN_RESET, TBN_CUSTHELP
                    Dim hWndFocus As LongPtr
                    hWndFocus = GetFocus()
                    Select Case NM.Code
                        Case TBN_RESET
                            Call ResetCustomizeButtons
                            Dim CloseDialog As Boolean
                            RaiseEvent ResetCustomizations(CloseDialog)
                            If GetFocus() <> hWndFocus And hWndFocus <> NULL_PTR Then SetFocusAPI hWndFocus
                            If CloseDialog = True Then
                                WindowProcUserControl = TBNRF_ENDCUSTOMIZE
                            Else
                                Call AllocCustomizeButtons
                            End If
                            Exit Function
                        Case TBN_CUSTHELP
                            RaiseEvent CustomizationHelp
                            If GetFocus() <> hWndFocus And hWndFocus <> NULL_PTR Then SetFocusAPI hWndFocus
                    End Select
                Case TBN_ENDADJUST
                    RaiseEvent EndCustomization
                Case TBN_GETBUTTONINFO
                    CopyMemory NMTB, ByVal lParam, LenB(NMTB)
                    If NMTB.iItem >= 0 And NMTB.iItem < ToolBarCustomizeButtonsCount Then
                        With ToolBarCustomizeButtons(NMTB.iItem + 1)
                        LSet NMTB.TBB = .TBB
                        NMTB.TBB.iString = StrPtr(.Caption)
                        If (NMTB.TBB.fsState And TBSTATE_CHECKED) = TBSTATE_CHECKED Then NMTB.TBB.fsState = NMTB.TBB.fsState And Not TBSTATE_CHECKED
                        If (NMTB.TBB.fsState And TBSTATE_PRESSED) = TBSTATE_PRESSED Then NMTB.TBB.fsState = NMTB.TBB.fsState And Not TBSTATE_PRESSED
                        If NMTB.TBB.dwData <> 0 Then
                            Set Button = PtrToObj(NMTB.TBB.dwData)
                            If Not Button.Description = vbNullString Then
                                Dim Description As String
                                Description = Left$(Button.Description, NMTB.cchText - 1) & vbNullChar
                                CopyMemory ByVal NMTB.pszText, ByVal StrPtr(Description), LenB(Description)
                            End If
                        End If
                        CopyMemory ByVal lParam, NMTB, LenB(NMTB)
                        WindowProcUserControl = 1
                        If .CX > 0 And (.TBB.fsStyle And BTNS_AUTOSIZE) = 0 Then PostMessage ToolBarHandle, UM_SETBUTTONCX, NMTB.iItem + 1, ByVal .CX
                        End With
                    Else
                        WindowProcUserControl = 0
                    End If
                    Exit Function
                Case NM_TOOLTIPSCREATED
                    Dim NMTTC As NMTOOLTIPSCREATED
                    CopyMemory NMTTC, ByVal lParam, LenB(NMTTC)
                    If NMTTC.hdr.hWndFrom = ToolBarHandle Then
                        ToolBarToolTipHandle = NMTTC.hWndToolTips
                        If ToolBarToolTipHandle <> NULL_PTR Then Call ComCtlsInitToolTip(ToolBarToolTipHandle)
                        Call SetVisualStylesToolTip
                    End If
            End Select
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function

Private Function WindowProcUserControlDesignMode(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
If wMsg = WM_NOTIFY Then
    Dim NM As NMHDR
    CopyMemory NM, ByVal lParam, LenB(NM)
    If NM.hWndFrom = ToolBarHandle And NM.Code = NM_CUSTOMDRAW Then
        WindowProcUserControlDesignMode = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
        Exit Function
    End If
End If
WindowProcUserControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
