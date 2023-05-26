VERSION 5.00
Begin VB.UserControl StatusBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "StatusBar.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "StatusBar.ctx":005E
   Begin VB.Timer TimerUpdatePanels 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
#If False Then
Private SbrStyleNormal, SbrStyleSimple
Private SbrPanelStyleText, SbrPanelStyleCaps, SbrPanelStyleNum, SbrPanelStyleIns, SbrPanelStyleScrl, SbrPanelStyleTime, SbrPanelStyleDate, SbrPanelStyleKana, SbrPanelStyleHangul, SbrPanelStyleJunja, SbrPanelStyleFinal, SbrPanelStyleKanji, SbrPanelStyleHanja
Private SbrPanelBevelFlat, SbrPanelBevelInset, SbrPanelBevelRaised
Private SbrPanelAutoSizeNone, SbrPanelAutoSizeSpring, SbrPanelAutoSizeContent
Private SbrPanelAlignmentLeft, SbrPanelAlignmentCenter, SbrPanelAlignmentRight
Private SbrPanelDTFormatShort, SbrPanelDTFormatLong
#End If
Public Enum SbrStyleConstants
SbrStyleNormal = 0
SbrStyleSimple = 1
End Enum
Public Enum SbrPanelStyleConstants
SbrPanelStyleText = 0
SbrPanelStyleCaps = 1
SbrPanelStyleNum = 2
SbrPanelStyleIns = 3
SbrPanelStyleScrl = 4
SbrPanelStyleTime = 5
SbrPanelStyleDate = 6
SbrPanelStyleKana = 7
SbrPanelStyleHangul = 8
SbrPanelStyleJunja = 9
SbrPanelStyleFinal = 10
SbrPanelStyleKanji = 11
SbrPanelStyleHanja = 12
End Enum
Public Enum SbrPanelBevelConstants
SbrPanelBevelFlat = 0
SbrPanelBevelInset = 1
SbrPanelBevelRaised = 2
End Enum
Public Enum SbrPanelAutoSizeConstants
SbrPanelAutoSizeNone = 0
SbrPanelAutoSizeSpring = 1
SbrPanelAutoSizeContent = 2
End Enum
Public Enum SbrPanelAlignmentConstants
SbrPanelAlignmentLeft = 0
SbrPanelAlignmentCenter = 1
SbrPanelAlignmentRight = 2
End Enum
Public Enum SbrPanelDTFormatConstants
SbrPanelDTFormatShort = 0
SbrPanelDTFormatLong = 1
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
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMMOUSE
hdr As NMHDR
dwItemSpec As Long
dwItemData As Long
PT As POINTAPI
dwHitInfo As Long
End Type
Private Type PAINTSTRUCT
hDC As Long
fErase As Long
RCPaint As RECT
fRestore As Long
fIncUpdate As Long
RGBReserved(0 To 31) As Byte
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
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event StyleChange()
Attribute StyleChange.VB_Description = "Occurs when the style changes."
Public Event PanelClick(ByVal Panel As SbrPanel, ByVal Button As Integer)
Attribute PanelClick.VB_Description = "Occurs when a user presses and then releases a mouse button over any of the panels."
Public Event PanelDblClick(ByVal Panel As SbrPanel, ByVal Button As Integer)
Attribute PanelDblClick.VB_Description = "Occurs when a user presses and then releases a mouse button twice over any of the panels."
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
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (ByRef pbKeyState As Byte) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal fMode As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lData As Long, ByVal wData As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal fFlags As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As SIZEAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function PtInRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Const ICC_BAR_CLASSES As Long = &H20
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const GWL_STYLE As Long = (-16)
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const TA_RTLREADING As Long = &H100
Private Const SM_CXVSCROLL As Long = 2
Private Const DST_TEXT As Long = &H1
Private Const DSS_DISABLED As Long = &H20
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_TOPMOST As Long = &H8
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
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
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_WINDOWPOSCHANGED As Long = &H47
Private Const WM_SIZE As Long = &H5
Private Const WM_DRAWITEM As Long = &H2B
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINT As Long = &H317, PRF_CLIENT As Long = &H4, PRF_ERASEBKGND As Long = &H8
Private Const SB_SIMPLEID As Long = &HFF
Private Const WM_USER As Long = &H400
Private Const SB_SETTEXTA As Long = (WM_USER + 1)
Private Const SB_SETTEXTW As Long = (WM_USER + 11)
Private Const SB_SETTEXT As Long = SB_SETTEXTW
Private Const SB_GETTEXTA As Long = (WM_USER + 2)
Private Const SB_GETTEXTW As Long = (WM_USER + 13)
Private Const SB_GETTEXT As Long = SB_GETTEXTW
Private Const SB_GETTEXTLENGTHA As Long = (WM_USER + 3)
Private Const SB_GETTEXTLENGTHW As Long = (WM_USER + 12)
Private Const SB_GETTEXTLENGTH As Long = SB_GETTEXTLENGTHW
Private Const SB_SETPARTS As Long = (WM_USER + 4)
Private Const SB_GETPARTS As Long = (WM_USER + 6)
Private Const SB_GETBORDERS As Long = (WM_USER + 7)
Private Const SB_SETMINHEIGHT As Long = (WM_USER + 8)
Private Const SB_SIMPLE As Long = (WM_USER + 9)
Private Const SB_GETRECT As Long = (WM_USER + 10)
Private Const SB_ISSIMPLE As Long = (WM_USER + 14)
Private Const TTM_UPDATE As Long = (WM_USER + 29)
Private Const TTM_ADDTOOLA As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW As Long = (WM_USER + 50)
Private Const TTM_ADDTOOL As Long = TTM_ADDTOOLW
Private Const TTM_DELTOOLA As Long = (WM_USER + 5)
Private Const TTM_DELTOOLW As Long = (WM_USER + 51)
Private Const TTM_DELTOOL As Long = TTM_DELTOOLW
Private Const TTM_NEWTOOLRECTA As Long = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW As Long = (WM_USER + 52)
Private Const TTM_NEWTOOLRECT As Long = TTM_NEWTOOLRECTW
Private Const TTM_GETTOOLINFOA As Long = (WM_USER + 8)
Private Const TTM_GETTOOLINFOW As Long = (WM_USER + 53)
Private Const TTM_GETTOOLINFO As Long = TTM_GETTOOLINFOW
Private Const TTM_SETTOOLINFOA As Long = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW As Long = (WM_USER + 54)
Private Const TTM_SETTOOLINFO As Long = TTM_SETTOOLINFOW
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTF_PARSELINKS As Long = &H1000
Private Const TTF_RTLREADING As Long = &H4
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX As Long = &H2
Private Const CCS_BOTTOM As Long = &H3
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETBKCOLOR As Long = (CCM_FIRST + 1)
Private Const SB_SETBKCOLOR As Long = CCM_SETBKCOLOR
Private Const SBT_OWNERDRAW As Long = &H1000
Private Const SBT_NOBORDERS As Long = &H100
Private Const SBT_POPOUT As Long = &H200
Private Const SBT_RTLREADING As Long = &H400 ' Useless on SBT_OWNERDRAW
Private Const SBT_TOOLTIPS As Long = &H800 ' Useless on SBT_OWNERDRAW
Private Const SBN_FIRST As Long = (-880)
Private Const SBN_SIMPLEMODECHANGE As Long = (SBN_FIRST - 0)
Private Const NM_FIRST As Long = 0
Private Const NM_CLICK As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK As Long = (NM_FIRST - 3)
Private Const NM_RCLICK As Long = (NM_FIRST - 5)
Private Const NM_RDBLCLK As Long = (NM_FIRST - 6)
Private Const SBARS_SIZEGRIP As Long = &H100
Private Const SBARS_TOOLTIPS As Long = SBT_TOOLTIPS ' Useless on SBT_OWNERDRAW
Private Const SBB_HORIZONTAL As Long = 0
Private Const SBB_VERTICAL As Long = 1
Private Const SBB_DIVIDER As Long = 2
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IPerPropertyBrowsingVB
Private Type InitPanelStruct
Text As String
Key As String
Tag As String
ToolTipText As String
Style As SbrPanelStyleConstants
Bevel As SbrPanelBevelConstants
AutoSize As SbrPanelAutoSizeConstants
Alignment As SbrPanelAlignmentConstants
DTFormat As SbrPanelDTFormatConstants
ForeColor As OLE_COLOR
MinWidth As Long
Picture As IPictureDisp
Enabled As Boolean
Visible As Boolean
Bold As Boolean
End Type
Private Type ShadowPanelStruct
Text As String
DisplayText As String
ToolTipText As String
ToolTipID As Long
Style As SbrPanelStyleConstants
Bevel As SbrPanelBevelConstants
AutoSize As SbrPanelAutoSizeConstants
Alignment As SbrPanelAlignmentConstants
DTFormat As SbrPanelDTFormatConstants
ForeColor As OLE_COLOR
MinWidth As Long
Picture As IPictureDisp
Enabled As Boolean
Visible As Boolean
Bold As Boolean
PictureRenderFlag As Integer
FixedWidth As Long
Left As Long
ActualWidth As Long
End Type
Private StatusBarHandle As Long, StatusBarToolTipHandle As Long
Private StatusBarSizeGripAllowable As Boolean
Private StatusBarParentForm As VB.Form
Private WithEvents StatusBarParentMDIFormEvents As VB.MDIForm
Attribute StatusBarParentMDIFormEvents.VB_VarHelpID = -1
Private WithEvents StatusBarParentFormEvents As VB.Form
Attribute StatusBarParentFormEvents.VB_VarHelpID = -1
Private StatusBarFontHandle As Long, StatusBarBoldFontHandle As Long
Private StatusBarIsClick As Boolean
Private StatusBarMouseOver As Boolean
Private StatusBarDesignMode As Boolean
Private StatusBarDoubleBufferEraseBkgDC As Long
Private StatusBarAlignable As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropPanels As SbrPanels
Private PropShadowPanelsCount As Long
Private PropShadowPanels() As ShadowPanelStruct
Private PropShadowDefaultPanel As ShadowPanelStruct
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropStyle As SbrStyleConstants
Private PropSimpleText As String
Private PropAllowSizeGrip As Boolean
Private PropShowTips As Boolean
Private PropBackColor As OLE_COLOR
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
Call ComCtlsInitCC(ICC_BAR_CLASSES)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
PropShadowDefaultPanel.FixedWidth = -1
End Sub

Private Sub UserControl_Show()
If StatusBarDesignMode = True Then
    Dim Align As Integer
    If StatusBarAlignable = True Then Align = Extender.Align Else Align = vbAlignNone
    If Align <> vbAlignBottom Then
        StatusBarSizeGripAllowable = False
        Call ReCreateStatusBar
    Else
        Call UserControl_Resize
    End If
End If
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then StatusBarAlignable = False Else StatusBarAlignable = True
StatusBarDesignMode = Not Ambient.UserMode
On Error GoTo 0
If StatusBarAlignable = True Then Extender.Align = vbAlignBottom
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropStyle = SbrStyleNormal
PropSimpleText = vbNullString
PropAllowSizeGrip = True
PropShowTips = False
PropBackColor = vbButtonFace
PropDoubleBuffer = True
If StatusBarAlignable = True Then StatusBarSizeGripAllowable = CBool((GetWindowLong(UserControl.ContainerHwnd, GWL_STYLE) And WS_THICKFRAME) = WS_THICKFRAME) Else StatusBarSizeGripAllowable = False
If StatusBarDesignMode = False Then
    On Error Resume Next
    With UserControl
    If .ParentControls.Count = 0 Then
    Else
        If TypeOf .Parent Is VB.MDIForm Then
            Set StatusBarParentForm = .Parent
            Set StatusBarParentMDIFormEvents = .Parent
        ElseIf TypeOf .Parent Is VB.Form Then
            Set StatusBarParentForm = .Parent
            Set StatusBarParentFormEvents = .Parent
        End If
    End If
    End With
    On Error GoTo 0
End If
Call CreateStatusBar
Me.Panels.Add
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then StatusBarAlignable = False Else StatusBarAlignable = True
StatusBarDesignMode = Not Ambient.UserMode
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
PropStyle = .ReadProperty("Style", SbrStyleNormal)
PropSimpleText = VarToStr(.ReadProperty("SimpleText", vbNullString))
PropAllowSizeGrip = .ReadProperty("AllowSizeGrip", True)
PropShowTips = .ReadProperty("ShowTips", False)
PropBackColor = .ReadProperty("BackColor", vbButtonFace)
PropDoubleBuffer = .ReadProperty("DoubleBuffer", True)
End With
With New PropertyBag
On Error Resume Next
.Contents = PropBag.ReadProperty("InitPanels", 0)
On Error GoTo 0
Dim InitPanelsCount As Long, i As Long
Dim InitPanels() As InitPanelStruct
InitPanelsCount = .ReadProperty("InitPanelsCount", 0)
If InitPanelsCount > 0 Then
    ReDim InitPanels(1 To InitPanelsCount) As InitPanelStruct
    For i = 1 To InitPanelsCount
        InitPanels(i).Text = VarToStr(.ReadProperty("InitPanelsText" & CStr(i), vbNullString))
        InitPanels(i).Key = VarToStr(.ReadProperty("InitPanelsKey" & CStr(i), vbNullString))
        InitPanels(i).Tag = VarToStr(.ReadProperty("InitPanelsTag" & CStr(i), vbNullString))
        InitPanels(i).ToolTipText = VarToStr(.ReadProperty("InitPanelsToolTipText" & CStr(i), vbNullString))
        InitPanels(i).Style = .ReadProperty("InitPanelsStyle" & CStr(i), SbrPanelStyleText)
        InitPanels(i).Bevel = .ReadProperty("InitPanelsBevel" & CStr(i), SbrPanelBevelInset)
        InitPanels(i).AutoSize = .ReadProperty("InitPanelsAutoSize" & CStr(i), SbrPanelAutoSizeNone)
        InitPanels(i).Alignment = .ReadProperty("InitPanelsAlignment" & CStr(i), SbrPanelAlignmentLeft)
        InitPanels(i).DTFormat = .ReadProperty("InitPanelsDTFormat" & CStr(i), SbrPanelDTFormatShort)
        InitPanels(i).ForeColor = .ReadProperty("InitPanelsForeColor" & CStr(i), vbButtonText)
        InitPanels(i).MinWidth = (.ReadProperty("InitPanelsMinWidth" & CStr(i), 96) * PixelsPerDIP_X())
        Set InitPanels(i).Picture = .ReadProperty("InitPanelsPicture" & CStr(i), Nothing)
        InitPanels(i).Enabled = .ReadProperty("InitPanelsEnabled" & CStr(i), True)
        InitPanels(i).Visible = .ReadProperty("InitPanelsVisible" & CStr(i), True)
        InitPanels(i).Bold = .ReadProperty("InitPanelsBold" & CStr(i), False)
    Next i
End If
End With
If StatusBarDesignMode = False Then
    On Error Resume Next
    With UserControl
    If .ParentControls.Count = 0 Then
    Else
        If TypeOf .Parent Is VB.MDIForm Then
            Set StatusBarParentForm = .Parent
            Set StatusBarParentMDIFormEvents = .Parent
        ElseIf TypeOf .Parent Is VB.Form Then
            Set StatusBarParentForm = .Parent
            Set StatusBarParentFormEvents = .Parent
        End If
    End If
    End With
    On Error GoTo 0
End If
If StatusBarAlignable = True Then StatusBarSizeGripAllowable = CBool((GetWindowLong(UserControl.ContainerHwnd, GWL_STYLE) And WS_THICKFRAME) = WS_THICKFRAME) Else StatusBarSizeGripAllowable = False
Call CreateStatusBar
If InitPanelsCount > 0 And StatusBarHandle <> 0 Then
    For i = 1 To InitPanelsCount
        Me.Panels.Add i, InitPanels(i).Key, InitPanels(i).Text, InitPanels(i).Style
        Me.Panels(i).Tag = InitPanels(i).Tag
        PropShadowPanels(i).ToolTipText = InitPanels(i).ToolTipText
        PropShadowPanels(i).Bevel = InitPanels(i).Bevel
        PropShadowPanels(i).AutoSize = InitPanels(i).AutoSize
        PropShadowPanels(i).Alignment = InitPanels(i).Alignment
        PropShadowPanels(i).DTFormat = InitPanels(i).DTFormat
        PropShadowPanels(i).ForeColor = InitPanels(i).ForeColor
        PropShadowPanels(i).MinWidth = InitPanels(i).MinWidth
        Set PropShadowPanels(i).Picture = InitPanels(i).Picture
        PropShadowPanels(i).Enabled = InitPanels(i).Enabled
        PropShadowPanels(i).Visible = InitPanels(i).Visible
        PropShadowPanels(i).Bold = InitPanels(i).Bold
    Next i
    Call SetParts
    Call SetPanels
End If
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
.WriteProperty "Style", PropStyle, SbrStyleNormal
.WriteProperty "SimpleText", StrToVar(PropSimpleText), vbNullString
.WriteProperty "AllowSizeGrip", PropAllowSizeGrip, True
.WriteProperty "ShowTips", PropShowTips, False
.WriteProperty "BackColor", PropBackColor, vbButtonFace
.WriteProperty "DoubleBuffer", PropDoubleBuffer, True
End With
Dim Count As Long
Count = Me.Panels.Count
With New PropertyBag
.WriteProperty "InitPanelsCount", Count, 0
If Count > 0 Then
    Dim i As Long
    For i = 1 To Count
        .WriteProperty "InitPanelsText" & CStr(i), StrToVar(Me.Panels(i).Text), vbNullString
        .WriteProperty "InitPanelsKey" & CStr(i), StrToVar(Me.Panels(i).Key), vbNullString
        .WriteProperty "InitPanelsTag" & CStr(i), StrToVar(Me.Panels(i).Tag), vbNullString
        .WriteProperty "InitPanelsToolTipText" & CStr(i), StrToVar(Me.Panels(i).ToolTipText), vbNullString
        .WriteProperty "InitPanelsStyle" & CStr(i), Me.Panels(i).Style, SbrPanelStyleText
        .WriteProperty "InitPanelsBevel" & CStr(i), Me.Panels(i).Bevel, SbrPanelBevelInset
        .WriteProperty "InitPanelsAutoSize" & CStr(i), Me.Panels(i).AutoSize, SbrPanelAutoSizeNone
        .WriteProperty "InitPanelsAlignment" & CStr(i), Me.Panels(i).Alignment, SbrPanelAlignmentLeft
        .WriteProperty "InitPanelsDTFormat" & CStr(i), Me.Panels(i).DTFormat, SbrPanelDTFormatShort
        .WriteProperty "InitPanelsForeColor" & CStr(i), Me.Panels(i).ForeColor, vbButtonFace
        .WriteProperty "InitPanelsMinWidth" & CStr(i), (CLng(UserControl.ScaleX(Me.Panels(i).MinWidth, vbContainerSize, vbPixels)) / PixelsPerDIP_X()), 96
        .WriteProperty "InitPanelsPicture" & CStr(i), PropShadowPanels(i).Picture, Nothing
        .WriteProperty "InitPanelsEnabled" & CStr(i), Me.Panels(i).Enabled, True
        .WriteProperty "InitPanelsVisible" & CStr(i), Me.Panels(i).Visible, True
        .WriteProperty "InitPanelsBold" & CStr(i), Me.Panels(i).Bold, False
    Next i
End If
PropBag.WriteProperty "InitPanels", .Contents, 0
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
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl.Extender
Dim Align As Integer
If StatusBarAlignable = True Then Align = .Align Else Align = vbAlignNone
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
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
Call SetMinHeight
If StatusBarHandle <> 0 Then MoveWindow StatusBarHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
Call SetParts
If PropShowTips = True Then Call UpdateToolTipRects
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyStatusBar
Call ComCtlsReleaseShellMod
End Sub

Private Sub StatusBarParentMDIFormEvents_Resize()
Call StatusBarParentFormEvents_Resize
End Sub

Private Sub StatusBarParentFormEvents_Resize()
Static LastWindowState As Integer
Dim CurrentWindowState As Integer
CurrentWindowState = StatusBarParentForm.WindowState
If CurrentWindowState = vbMaximized Then
    If StatusBarSizeGripAllowable = True And Me.IncludesSizeGrip = True Then
        StatusBarSizeGripAllowable = False
        Call ReCreateStatusBar
    End If
ElseIf CurrentWindowState = vbNormal And LastWindowState = vbMaximized Then
    If StatusBarAlignable = True Then StatusBarSizeGripAllowable = CBool(((GetWindowLong(UserControl.ContainerHwnd, GWL_STYLE) And WS_THICKFRAME) = WS_THICKFRAME) And Extender.Align = vbAlignBottom) Else StatusBarSizeGripAllowable = False
    If StatusBarSizeGripAllowable = True And PropAllowSizeGrip = True And Me.IncludesSizeGrip = False Then
        Call ReCreateStatusBar
    End If
End If
LastWindowState = CurrentWindowState
End Sub

Private Sub TimerUpdatePanels_Timer()
If StatusBarHandle = 0 Then Exit Sub
Dim NeedUpdate As Boolean
If PropShadowPanelsCount > 0 Then
    Dim i As Long, Text As String, Enabled As Boolean, RC As RECT
    For i = 1 To PropShadowPanelsCount
        With PropShadowPanels(i)
        If .Visible = True Then
            Call GetDisplayText(i, Text, Enabled)
            If StrComp(Text, .DisplayText) <> 0 Then
                InvalidateRect StatusBarHandle, ByVal 0&, 1
                Call SetParts
                NeedUpdate = True
                Exit For
            ElseIf Enabled Xor .Enabled Then
                Call GetPanelRect(i, RC)
                InvalidateRect StatusBarHandle, ByVal VarPtr(RC), 1
                NeedUpdate = True
            End If
        End If
        End With
    Next i
End If
If NeedUpdate = True Then UpdateWindow StatusBarHandle
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
hWnd = StatusBarHandle
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
Dim OldFontHandle As Long, OldBoldFontHandle As Long
Dim TempFont As StdFont
Set PropFont = NewFont
OldFontHandle = StatusBarFontHandle
OldBoldFontHandle = StatusBarBoldFontHandle
StatusBarFontHandle = CreateGDIFontFromOLEFont(PropFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Bold = True
StatusBarBoldFontHandle = CreateGDIFontFromOLEFont(TempFont)
If StatusBarHandle <> 0 Then SendMessage StatusBarHandle, WM_SETFONT, StatusBarFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If OldBoldFontHandle <> 0 Then DeleteObject OldBoldFontHandle
Call SetMinHeight
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long, OldBoldFontHandle As Long
Dim TempFont As StdFont
OldFontHandle = StatusBarFontHandle
OldBoldFontHandle = StatusBarBoldFontHandle
StatusBarFontHandle = CreateGDIFontFromOLEFont(PropFont)
Set TempFont = CloneOLEFont(PropFont)
TempFont.Bold = True
StatusBarBoldFontHandle = CreateGDIFontFromOLEFont(TempFont)
If StatusBarHandle <> 0 Then SendMessage StatusBarHandle, WM_SETFONT, StatusBarFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
If OldBoldFontHandle <> 0 Then DeleteObject OldBoldFontHandle
Call SetMinHeight
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If StatusBarHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles StatusBarHandle
    Else
        RemoveVisualStyles StatusBarHandle
    End If
    Call SetVisualStylesToolTip
End If
If PropVisualStyles = False Or EnabledVisualStyles() = False Then
    UserControl.BackColor = PropBackColor
Else
    UserControl.BackColor = vbButtonFace
End If
Me.Refresh
UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If StatusBarHandle <> 0 Then EnableWindow StatusBarHandle, IIf(Value = True, 1, 0)
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
If StatusBarDesignMode = False Then Call RefreshMousePointer
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
        If StatusBarDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If StatusBarDesignMode = False Then Call RefreshMousePointer
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
If StatusBarDesignMode = False Then Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
If StatusBarHandle <> 0 Then Call ComCtlsSetRightToLeft(StatusBarHandle, dwMask)
Me.SimpleText = Me.SimpleText
If StatusBarToolTipHandle <> 0 Then
    If PropRightToLeft = True Then
        If PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = WS_EX_RTLREADING
    Else
        dwMask = 0
    End If
    Call ComCtlsSetRightToLeft(StatusBarToolTipHandle, dwMask)
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

Public Property Get Style() As SbrStyleConstants
Attribute Style.VB_Description = "Returns/sets the single (simple) or multiple panel (normal) style."
If StatusBarHandle <> 0 Then
    If SendMessage(StatusBarHandle, SB_ISSIMPLE, 0, ByVal 0&) = 0 Then
        Style = SbrStyleNormal
    Else
        Style = SbrStyleSimple
    End If
Else
    Style = PropStyle
End If
End Property

Public Property Let Style(ByVal Value As SbrStyleConstants)
Select Case Value
    Case SbrStyleNormal, SbrStyleSimple
        PropStyle = Value
    Case Else
        Err.Raise 380
End Select
If StatusBarHandle <> 0 Then SendMessage StatusBarHandle, SB_SIMPLE, IIf(PropStyle = SbrStyleSimple, 1, 0), ByVal 0&
UserControl.PropertyChanged "Style"
End Property

Public Property Get SimpleText() As String
Attribute SimpleText.VB_Description = "Returns/sets the text displayed when the style property is set to simple."
Attribute SimpleText.VB_MemberFlags = "200"
If StatusBarHandle <> 0 Then
    If SendMessage(StatusBarHandle, SB_ISSIMPLE, 0, ByVal 0&) <> 0 Then
        Dim Length As Long
        Length = CIntToUInt(LoWord(SendMessage(StatusBarHandle, SB_GETTEXTLENGTH, 0, ByVal 0&)))
        If Length > 0 Then
            SimpleText = String$(Length, vbNullChar)
            SendMessage StatusBarHandle, SB_GETTEXT, 0, ByVal StrPtr(SimpleText)
        End If
    Else
        SimpleText = PropSimpleText
    End If
Else
    SimpleText = PropSimpleText
End If
End Property

Public Property Let SimpleText(ByVal Value As String)
PropSimpleText = Value
If StatusBarHandle <> 0 Then
    Dim Style As Long
    Style = 0
    If PropRightToLeft = True And PropRightToLeftLayout = False Then Style = Style Or SBT_RTLREADING
    SendMessage StatusBarHandle, SB_SETTEXT, SB_SIMPLEID Or Style, ByVal StrPtr(PropSimpleText)
End If
UserControl.PropertyChanged "SimpleText"
End Property

Public Property Get AllowSizeGrip() As Boolean
Attribute AllowSizeGrip.VB_Description = "Returns/sets a value indicating if the control is allowed to include a size grip at the right end."
AllowSizeGrip = PropAllowSizeGrip
End Property

Public Property Let AllowSizeGrip(ByVal Value As Boolean)
PropAllowSizeGrip = Value
If StatusBarHandle <> 0 Then Call ReCreateStatusBar
UserControl.PropertyChanged "AllowSizeGrip"
End Property

Public Property Get ShowTips() As Boolean
Attribute ShowTips.VB_Description = "Returns/sets a value that determines whether the tool tip text properties will be displayed or not."
ShowTips = PropShowTips
End Property

Public Property Let ShowTips(ByVal Value As Boolean)
PropShowTips = Value
If StatusBarHandle <> 0 And StatusBarDesignMode = False Then
    If PropShowTips = False Then
        Call DestroyToolTip
    Else
        Call CreateToolTip
    End If
    Call SetPanels
End If
UserControl.PropertyChanged "ShowTips"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object. This property is ignored if the version of comctl32.dll is 6.0 or higher and the visual styles property is set to true."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If StatusBarHandle <> 0 Then SendMessage StatusBarHandle, SB_SETBKCOLOR, 0, ByVal WinColor(PropBackColor)
If PropVisualStyles = False Or EnabledVisualStyles() = False Then
    UserControl.BackColor = PropBackColor
Else
    UserControl.BackColor = vbButtonFace
End If
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get DoubleBuffer() As Boolean
Attribute DoubleBuffer.VB_Description = "Returns/sets a value that determines whether the control paints via double-buffering, which reduces flicker."
DoubleBuffer = PropDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal Value As Boolean)
PropDoubleBuffer = Value
UserControl.PropertyChanged "DoubleBuffer"
End Property

Public Property Get Panels() As SbrPanels
Attribute Panels.VB_Description = "Returns a reference to a collection of the panel objects."
If PropPanels Is Nothing Then
    Set PropPanels = New SbrPanels
    PropPanels.FInit Me
End If
Set Panels = PropPanels
End Property

Friend Sub FPanelsAdd(ByVal Index As Long, Optional ByVal Text As String, Optional ByVal Style As SbrPanelStyleConstants)
PropShadowPanelsCount = PropShadowPanelsCount + 1
Dim PanelIndex As Long
If Index = 0 Then
    PanelIndex = PropShadowPanelsCount
Else
    PanelIndex = Index
End If
ReDim Preserve PropShadowPanels(1 To PropShadowPanelsCount) As ShadowPanelStruct
Dim i As Long
If PanelIndex < PropShadowPanelsCount Then
    For i = PropShadowPanelsCount To PanelIndex + 1 Step -1
        LSet PropShadowPanels(i) = PropShadowPanels(i - 1)
    Next i
End If
LSet PropShadowPanels(PanelIndex) = PropShadowDefaultPanel
With PropShadowPanels(PanelIndex)
.Text = Text
.ToolTipText = vbNullString
.ToolTipID = NextToolTipID()
.MinWidth = (96 * PixelsPerDIP_X())
Select Case Style
    Case SbrPanelStyleText, SbrPanelStyleCaps, SbrPanelStyleNum, SbrPanelStyleIns, SbrPanelStyleScrl, SbrPanelStyleTime, SbrPanelStyleDate, SbrPanelStyleKana, SbrPanelStyleHangul, SbrPanelStyleJunja, SbrPanelStyleFinal, SbrPanelStyleKanji, SbrPanelStyleHanja
        .Style = Style
    Case Else
        Err.Raise 380
End Select
.Bevel = SbrPanelBevelInset
.AutoSize = SbrPanelAutoSizeNone
.Alignment = SbrPanelAlignmentLeft
.ForeColor = vbButtonText
Set .Picture = Nothing
.Enabled = True
.Visible = True
.Bold = False
Call GetDisplayText(PanelIndex, .DisplayText)
End With
Call SetParts
Call SetPanels
Call CheckTimer
UserControl.PropertyChanged "InitPanels"
End Sub

Friend Sub FPanelsRemove(ByVal Index As Long)
If PropShowTips = True Then
    If StatusBarHandle <> 0 And StatusBarToolTipHandle <> 0 Then
        Dim TI As TOOLINFO
        With TI
        .cbSize = LenB(TI)
        .hWnd = StatusBarHandle
        .uId = PropShadowPanels(Index).ToolTipID
        End With
        SendMessage StatusBarToolTipHandle, TTM_DELTOOL, 0, ByVal VarPtr(TI)
    End If
End If
Dim i As Long
For i = Index To PropShadowPanelsCount - 1
    LSet PropShadowPanels(i) = PropShadowPanels(i + 1)
Next i
PropShadowPanelsCount = PropShadowPanelsCount - 1
If PropShadowPanelsCount > 0 Then
    ReDim Preserve PropShadowPanels(1 To PropShadowPanelsCount) As ShadowPanelStruct
Else
    Erase PropShadowPanels()
End If
Call SetParts
If PropShadowPanelsCount > 0 Then Call SetPanels
Call CheckTimer
UserControl.PropertyChanged "InitPanels"
End Sub

Friend Sub FPanelsClear()
Dim i As Long
For i = 1 To PropShadowPanelsCount
    Me.FPanelsRemove 1
Next i
End Sub

Friend Property Get FPanelText(ByVal Index As Long) As String
If StatusBarHandle <> 0 Then FPanelText = PropShadowPanels(Index).Text
End Property

Friend Property Let FPanelText(ByVal Index As Long, ByVal Value As String)
If PropShadowPanels(Index).Text = Value Then Exit Property
If StatusBarHandle <> 0 Then
    PropShadowPanels(Index).Text = Value
    Call SetPanelText(Index)
    Call GetDisplayText(Index, PropShadowPanels(Index).DisplayText)
    If PropShadowPanels(Index).AutoSize = SbrPanelAutoSizeContent Then Call SetParts
End If
End Property

Friend Property Get FPanelToolTipText(ByVal Index As Long) As String
If StatusBarHandle <> 0 Then FPanelToolTipText = PropShadowPanels(Index).ToolTipText
End Property

Friend Property Let FPanelToolTipText(ByVal Index As Long, ByVal Value As String)
If StatusBarHandle <> 0 Then
    PropShadowPanels(Index).ToolTipText = Value
    If PropShowTips = True Then Call SetPanelToolTipText(Index)
End If
End Property

Friend Property Get FPanelStyle(ByVal Index As Long) As SbrPanelStyleConstants
If StatusBarHandle <> 0 Then FPanelStyle = PropShadowPanels(Index).Style
End Property

Friend Property Let FPanelStyle(ByVal Index As Long, ByVal Value As SbrPanelStyleConstants)
If StatusBarHandle <> 0 Then
    Select Case Value
        Case SbrPanelStyleText, SbrPanelStyleCaps, SbrPanelStyleNum, SbrPanelStyleIns, SbrPanelStyleScrl, SbrPanelStyleTime, SbrPanelStyleDate, SbrPanelStyleKana, SbrPanelStyleHangul, SbrPanelStyleJunja, SbrPanelStyleFinal, SbrPanelStyleKanji, SbrPanelStyleHanja
            PropShadowPanels(Index).Style = Value
        Case Else
            Err.Raise 380
    End Select
    Dim RC As RECT
    Call GetPanelRect(Index, RC)
    InvalidateRect StatusBarHandle, ByVal VarPtr(RC), 1
    UpdateWindow StatusBarHandle
    If PropShadowPanels(Index).AutoSize = SbrPanelAutoSizeContent Then Call SetParts
End If
End Property

Friend Property Get FPanelBevel(ByVal Index As Long) As SbrPanelBevelConstants
If StatusBarHandle <> 0 Then FPanelBevel = PropShadowPanels(Index).Bevel
End Property

Friend Property Let FPanelBevel(ByVal Index As Long, ByVal Value As SbrPanelBevelConstants)
If StatusBarHandle <> 0 Then
    Select Case Value
        Case SbrPanelBevelFlat, SbrPanelBevelInset, SbrPanelBevelRaised
            PropShadowPanels(Index).Bevel = Value
        Case Else
            Err.Raise 380
    End Select
    Call SetPanelText(Index)
    Me.Refresh
End If
End Property

Friend Property Get FPanelAutoSize(ByVal Index As Long) As SbrPanelAutoSizeConstants
If StatusBarHandle <> 0 Then FPanelAutoSize = PropShadowPanels(Index).AutoSize
End Property

Friend Property Let FPanelAutoSize(ByVal Index As Long, ByVal Value As SbrPanelAutoSizeConstants)
If StatusBarHandle <> 0 Then
    Select Case Value
        Case SbrPanelAutoSizeNone, SbrPanelAutoSizeSpring, SbrPanelAutoSizeContent
            PropShadowPanels(Index).AutoSize = Value
        Case Else
            Err.Raise 380
    End Select
    Call SetParts
    Call SetPanels
End If
End Property

Friend Property Get FPanelAlignment(ByVal Index As Long) As SbrPanelAlignmentConstants
If StatusBarHandle <> 0 Then FPanelAlignment = PropShadowPanels(Index).Alignment
End Property

Friend Property Let FPanelAlignment(ByVal Index As Long, ByVal Value As SbrPanelAlignmentConstants)
If StatusBarHandle <> 0 Then
    Select Case Value
        Case SbrPanelAlignmentLeft, SbrPanelAlignmentCenter, SbrPanelAlignmentRight
            PropShadowPanels(Index).Alignment = Value
        Case Else
            Err.Raise 380
    End Select
    Dim RC As RECT
    Call GetPanelRect(Index, RC)
    InvalidateRect StatusBarHandle, ByVal VarPtr(RC), 1
    UpdateWindow StatusBarHandle
End If
End Property

Friend Property Get FPanelDTFormat(ByVal Index As Long) As SbrPanelDTFormatConstants
If StatusBarHandle <> 0 Then FPanelDTFormat = PropShadowPanels(Index).DTFormat
End Property

Friend Property Let FPanelDTFormat(ByVal Index As Long, ByVal Value As SbrPanelDTFormatConstants)
If StatusBarHandle <> 0 Then
    Select Case Value
        Case SbrPanelDTFormatShort, SbrPanelDTFormatLong
            PropShadowPanels(Index).DTFormat = Value
        Case Else
            Err.Raise 380
    End Select
    Dim RC As RECT
    Call GetPanelRect(Index, RC)
    InvalidateRect StatusBarHandle, ByVal VarPtr(RC), 1
    UpdateWindow StatusBarHandle
    If PropShadowPanels(Index).AutoSize = SbrPanelAutoSizeContent Then
        Select Case PropShadowPanels(Index).Style
            Case SbrPanelStyleTime, SbrPanelStyleDate
                Call SetParts
        End Select
    End If
End If
End Property

Friend Property Get FPanelForeColor(ByVal Index As Long) As OLE_COLOR
If StatusBarHandle <> 0 Then FPanelForeColor = PropShadowPanels(Index).ForeColor
End Property

Friend Property Let FPanelForeColor(ByVal Index As Long, ByVal Value As OLE_COLOR)
If StatusBarHandle <> 0 Then
    PropShadowPanels(Index).ForeColor = Value
    Dim RC As RECT
    Call GetPanelRect(Index, RC)
    InvalidateRect StatusBarHandle, ByVal VarPtr(RC), 1
    UpdateWindow StatusBarHandle
End If
End Property

Friend Property Get FPanelMinWidth(ByVal Index As Long) As Single
If StatusBarHandle <> 0 Then FPanelMinWidth = UserControl.ScaleX(PropShadowPanels(Index).MinWidth, vbPixels, vbContainerSize)
End Property

Friend Property Let FPanelMinWidth(ByVal Index As Long, ByVal Value As Single)
If Value < 0 Then Err.Raise 380
If StatusBarHandle <> 0 Then
    PropShadowPanels(Index).MinWidth = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    Call SetParts
    Call SetPanels
End If
End Property

Friend Property Get FPanelPicture(ByVal Index As Long) As IPictureDisp
If StatusBarHandle <> 0 Then Set FPanelPicture = PropShadowPanels(Index).Picture
End Property

Friend Property Let FPanelPicture(ByVal Index As Long, ByVal Value As IPictureDisp)
Set Me.FPanelPicture(Index) = Value
End Property

Friend Property Set FPanelPicture(ByVal Index As Long, ByVal Value As IPictureDisp)
If StatusBarHandle <> 0 Then
    Set PropShadowPanels(Index).Picture = Value
    PropShadowPanels(Index).PictureRenderFlag = 0
    Call SetParts
    Call SetPanels
End If
End Property

Friend Property Get FPanelEnabled(ByVal Index As Long) As Boolean
If StatusBarHandle <> 0 Then FPanelEnabled = PropShadowPanels(Index).Enabled
End Property

Friend Property Let FPanelEnabled(ByVal Index As Long, ByVal Value As Boolean)
If StatusBarHandle <> 0 Then
    PropShadowPanels(Index).Enabled = Value
    Dim RC As RECT
    Call GetPanelRect(Index, RC)
    InvalidateRect StatusBarHandle, ByVal VarPtr(RC), 1
    UpdateWindow StatusBarHandle
End If
End Property

Friend Property Get FPanelVisible(ByVal Index As Long) As Boolean
If StatusBarHandle <> 0 Then FPanelVisible = PropShadowPanels(Index).Visible
End Property

Friend Property Let FPanelVisible(ByVal Index As Long, ByVal Value As Boolean)
If StatusBarHandle <> 0 Then
    PropShadowPanels(Index).Visible = Value
    Call SetParts
    Call SetPanels
End If
End Property

Friend Property Get FPanelBold(ByVal Index As Long) As Boolean
If StatusBarHandle <> 0 Then FPanelBold = PropShadowPanels(Index).Bold
End Property

Friend Property Let FPanelBold(ByVal Index As Long, ByVal Value As Boolean)
If StatusBarHandle <> 0 Then
    PropShadowPanels(Index).Bold = Value
    Dim RC As RECT
    Call GetPanelRect(Index, RC)
    InvalidateRect StatusBarHandle, ByVal VarPtr(RC), 1
    UpdateWindow StatusBarHandle
End If
End Property

Friend Property Get FPanelLeft(ByVal Index As Long) As Single
If StatusBarHandle <> 0 Then FPanelLeft = UserControl.ScaleX(PropShadowPanels(Index).Left, vbPixels, vbContainerSize)
End Property

Friend Property Get FPanelWidth(ByVal Index As Long) As Single
If StatusBarHandle <> 0 Then FPanelWidth = UserControl.ScaleX(PropShadowPanels(Index).ActualWidth, vbPixels, vbContainerSize)
End Property

Friend Property Let FPanelWidth(ByVal Index As Long, ByVal Value As Single)
If Value < 0 Then
    If Value = -1 And PropShadowPanels(Index).AutoSize = SbrPanelAutoSizeNone Then
    Else
        Err.Raise 380
    End If
End If
If StatusBarHandle <> 0 Then
    If PropShadowPanels(Index).AutoSize <> SbrPanelAutoSizeSpring Then
        Select Case PropShadowPanels(Index).AutoSize
            Case SbrPanelAutoSizeNone
                If Value > -1 Then
                    PropShadowPanels(Index).FixedWidth = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
                Else
                    PropShadowPanels(Index).FixedWidth = -1
                End If
            Case SbrPanelAutoSizeContent
                PropShadowPanels(Index).MinWidth = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
        End Select
        Call SetParts
        Call SetPanels
    End If
End If
End Property

Private Sub CreateStatusBar()
If StatusBarHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS Or CCS_BOTTOM
If StatusBarSizeGripAllowable = True And PropAllowSizeGrip = True Then dwStyle = dwStyle Or SBARS_SIZEGRIP
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If StatusBarDesignMode = False Then
    ' The WM_NOTIFYFORMAT notification must be handled, which will be sent on control creation.
    ' Thus it is necessary to subclass the parent before the control is created.
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
StatusBarHandle = CreateWindowEx(dwExStyle, StrPtr("msctls_statusbar32"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Style = PropStyle
Me.SimpleText = PropSimpleText
Me.ShowTips = PropShowTips
Me.BackColor = PropBackColor
If StatusBarDesignMode = False Then
    If StatusBarHandle <> 0 Then Call ComCtlsSetSubclass(StatusBarHandle, Me, 1)
Else
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 3)
End If
Call SetMinHeight
Call CheckTimer
End Sub

Private Sub CreateToolTip()
Static Done As Boolean
Dim dwExStyle As Long
If StatusBarToolTipHandle <> 0 Then Exit Sub
If Done = False Then
    Call ComCtlsInitCC(ICC_TAB_CLASSES)
    Done = True
End If
dwExStyle = WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
StatusBarToolTipHandle = CreateWindowEx(dwExStyle, StrPtr("tooltips_class32"), StrPtr("Tool Tip"), WS_POPUP Or TTS_ALWAYSTIP Or TTS_NOPREFIX, 0, 0, 0, 0, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If StatusBarToolTipHandle <> 0 Then Call ComCtlsInitToolTip(StatusBarToolTipHandle)
Call SetVisualStylesToolTip
End Sub

Private Sub ReCreateStatusBar()
Dim Locked As Boolean
With Me
Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
Call DestroyStatusBar
Call CreateStatusBar
Call UserControl_Resize
Call SetParts
Call SetPanels
If Locked = True Then LockWindowUpdate 0
.Refresh
End With
End Sub

Private Sub DestroyStatusBar()
If StatusBarHandle = 0 Then Exit Sub
TimerUpdatePanels.Enabled = False
Call ComCtlsRemoveSubclass(StatusBarHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
Call DestroyToolTip
ShowWindow StatusBarHandle, SW_HIDE
SetParent StatusBarHandle, 0
DestroyWindow StatusBarHandle
StatusBarHandle = 0
If StatusBarFontHandle <> 0 Then
    DeleteObject StatusBarFontHandle
    StatusBarFontHandle = 0
End If
If StatusBarBoldFontHandle <> 0 Then
    DeleteObject StatusBarBoldFontHandle
    StatusBarBoldFontHandle = 0
End If
End Sub

Private Sub DestroyToolTip()
If StatusBarToolTipHandle = 0 Then Exit Sub
DestroyWindow StatusBarToolTipHandle
StatusBarToolTipHandle = 0
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Function IncludesSizeGrip() As Boolean
Attribute IncludesSizeGrip.VB_Description = "Returns a value indicating if the control includes a size grip at the right end."
If StatusBarHandle <> 0 Then IncludesSizeGrip = CBool((GetWindowLong(StatusBarHandle, GWL_STYLE) And SBARS_SIZEGRIP) = SBARS_SIZEGRIP)
End Function

Public Function HitTest(ByVal X As Single, ByVal Y As Single) As SbrPanel
Attribute HitTest.VB_Description = "Returns a reference to the panel item object located at the coordinates of X and Y."
If StatusBarHandle <> 0 Then
    Dim P As POINTAPI, RC As RECT, i As Long
    P.X = UserControl.ScaleX(X, vbContainerPosition, vbPixels)
    P.Y = UserControl.ScaleY(Y, vbContainerPosition, vbPixels)
    For i = 1 To PropShadowPanelsCount
        Call GetPanelRect(i, RC)
        If PtInRect(RC, P.X, P.Y) <> 0 Then
            Set HitTest = Me.Panels(i)
            Exit For
        End If
    Next i
End If
End Function

Private Sub GetDisplayText(ByVal Index As Long, ByRef Text As String, Optional ByRef Enabled As Boolean)
Static KeyState(0 To 255) As Byte
Text = vbNullString
Select Case PropShadowPanels(Index).Style
    Case SbrPanelStyleText
        Text = PropShadowPanels(Index).Text
        Enabled = PropShadowPanels(Index).Enabled
    Case SbrPanelStyleCaps
        Text = "CAPS"
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyCapital))
    Case SbrPanelStyleNum
        Text = "NUM"
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyNumlock))
    Case SbrPanelStyleIns
        Text = "INS"
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyInsert))
    Case SbrPanelStyleScrl
        Text = "SCRL"
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyScrollLock))
    Case SbrPanelStyleTime
        Select Case PropShadowPanels(Index).DTFormat
            Case SbrPanelDTFormatShort
                Text = VBA.FormatDateTime(VBA.Time, vbShortTime)
            Case SbrPanelDTFormatLong
                Text = VBA.FormatDateTime(VBA.Time, vbLongTime)
        End Select
        Enabled = PropShadowPanels(Index).Enabled
    Case SbrPanelStyleDate
        Select Case PropShadowPanels(Index).DTFormat
            Case SbrPanelDTFormatShort
                Text = VBA.FormatDateTime(VBA.Date, vbShortDate)
            Case SbrPanelDTFormatLong
                Text = VBA.FormatDateTime(VBA.Date, vbLongDate)
        End Select
        Enabled = PropShadowPanels(Index).Enabled
    Case SbrPanelStyleKana
        Text = "KANA"
        Const vbKeyKana As Long = &H15
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyKana))
    Case SbrPanelStyleHangul
        Text = "HANGUL"
        Const vbKeyHangul As Long = &H15
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyHangul))
    Case SbrPanelStyleJunja
        Text = "JUNJA"
        Const vbKeyJunja As Long = &H17
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyJunja))
    Case SbrPanelStyleFinal
        Text = "FINAL"
        Const vbKeyFinal As Long = &H18
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyFinal))
    Case SbrPanelStyleKanji
        Text = "KANJI"
        Const vbKeyKanji As Long = &H19
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyKanji))
    Case SbrPanelStyleHanja
        Text = "HANJA"
        Const vbKeyHanja As Long = &H19
        GetKeyboardState KeyState(0)
        Enabled = CBool(KeyState(vbKeyHanja))
End Select
End Sub

Private Sub DrawPanel(ByVal Index As Long, ByVal hDC As Long, ByRef RC As RECT)
If Index <> SB_SIMPLEID And StatusBarHandle <> 0 Then
    Dim Text As String, Size As SIZEAPI, OldTextAlign As Long, OldBkMode As Long, OldTextColor As Long, hFontOld As Long
    With PropShadowPanels(Index)
    Call GetDisplayText(Index, Text, .Enabled)
    .DisplayText = Text
    If StrPtr(Text) = 0 Then Text = ""
    OldBkMode = SetBkMode(hDC, 1)
    OldTextColor = SetTextColor(hDC, WinColor(.ForeColor))
    If .Bold = True And StatusBarBoldFontHandle <> 0 Then hFontOld = SelectObject(hDC, StatusBarBoldFontHandle)
    GetTextExtentPoint32 hDC, ByVal StrPtr(Text), Len(Text), Size
    Dim PictureWidth As Long, PictureHeight As Long
    Dim PictureLeft As Long, PictureTop As Long
    If Not .Picture Is Nothing Then
        If .Picture.Handle <> 0 Then
            PictureWidth = CHimetricToPixel_X(.Picture.Width)
            PictureHeight = CHimetricToPixel_Y(.Picture.Height)
            PictureTop = RC.Top + ((RC.Bottom - RC.Top) \ 2) - (PictureHeight \ 2)
        End If
    End If
    RC.Top = RC.Top + ((RC.Bottom - RC.Top) \ 2) - (Size.CY \ 2) - (1 * PixelsPerDIP_Y())
    Select Case .Alignment
        Case SbrPanelAlignmentLeft
            RC.Right = RC.Right + PictureWidth + (IIf(PictureWidth > 0, 5, 1) * PixelsPerDIP_X())
            RC.Left = RC.Left + PictureWidth + (IIf(PictureWidth > 0, 5, 1) * PixelsPerDIP_X())
        Case SbrPanelAlignmentCenter
            RC.Right = RC.Right - (1 * PixelsPerDIP_X())
            RC.Left = RC.Left + (((RC.Right - RC.Left) - (Size.CX - PictureWidth - (IIf(PictureWidth > 0, 4, 0) * PixelsPerDIP_X()))) / 2)
        Case SbrPanelAlignmentRight
            RC.Right = RC.Right - (1 * PixelsPerDIP_X())
            RC.Left = RC.Left + ((RC.Right - RC.Left) - Size.CX)
    End Select
    If PictureWidth > 0 And PictureHeight > 0 Then
        PictureLeft = RC.Left - (PictureWidth + (4 * PixelsPerDIP_X()))
        Call RenderPicture(.Picture, hDC, PictureLeft, PictureTop, PictureWidth, PictureHeight, .PictureRenderFlag)
    End If
    Dim Flags As Long
    Flags = DST_TEXT
    If .Enabled = False Then Flags = Flags Or DSS_DISABLED
    If PropRightToLeft = True And PropRightToLeftLayout = False Then OldTextAlign = SetTextAlign(hDC, TA_RTLREADING)
    DrawState hDC, 0, 0, StrPtr(Text), Len(Text), RC.Left, RC.Top, RC.Right - RC.Left, RC.Bottom - RC.Top, Flags
    If PropRightToLeft = True And PropRightToLeftLayout = False Then SetTextAlign hDC, OldTextAlign
    SetBkMode hDC, OldBkMode
    SetTextColor hDC, OldTextColor
    If hFontOld <> 0 Then SelectObject hDC, hFontOld
    End With
End If
End Sub

Private Sub SetMinHeight()
If StatusBarHandle <> 0 Then
    Dim Borders(0 To 2) As Long
    SendMessage StatusBarHandle, SB_GETBORDERS, 0, ByVal VarPtr(Borders(0))
    With UserControl
    Dim Height As Long, FontHeight As Long
    Dim hDC As Long
    Dim hFontOld As Long
    Dim Size As SIZEAPI
    Height = UserControl.ScaleHeight - Borders(SBB_VERTICAL)
    If StatusBarFontHandle <> 0 Then
        hDC = GetDC(StatusBarHandle)
        If hDC <> 0 Then
            hFontOld = SelectObject(hDC, StatusBarFontHandle)
            If hFontOld <> 0 Then
                GetTextExtentPoint32 hDC, StrPtr("A"), 1, Size
                FontHeight = Size.CY + Borders(SBB_VERTICAL)
                SelectObject hDC, hFontOld
            End If
            ReleaseDC StatusBarHandle, hDC
        End If
    End If
    If Height < FontHeight Then Height = FontHeight
    SendMessage StatusBarHandle, SB_SETMINHEIGHT, Height, ByVal 0&
    SendMessage StatusBarHandle, WM_SIZE, 0, ByVal 0&
    Dim WndRect As RECT
    GetWindowRect StatusBarHandle, WndRect
    On Error Resume Next
    .Extender.Height = .ScaleY((WndRect.Bottom - WndRect.Top), vbPixels, vbContainerSize)
    On Error GoTo 0
    MoveWindow StatusBarHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    End With
End If
End Sub

Private Sub SetParts()
If StatusBarHandle <> 0 Then
    Dim Parts() As Long
    If PropShadowPanelsCount > 0 Then
        Dim Borders(0 To 2) As Long
        Dim i As Long
        Dim TotalWidth As Long
        SendMessage StatusBarHandle, SB_GETBORDERS, 0, ByVal VarPtr(Borders(0))
        ReDim Parts(0 To PropShadowPanelsCount - 1) As Long
        For i = 1 To PropShadowPanelsCount
            Parts(i - 1) = GetGoodWidth(i)
            TotalWidth = TotalWidth + Parts(i - 1)
            If i < PropShadowPanelsCount Then TotalWidth = TotalWidth + Borders(SBB_DIVIDER)
        Next i
        TotalWidth = TotalWidth + Borders(SBB_HORIZONTAL) + Borders(SBB_HORIZONTAL)
        If Me.IncludesSizeGrip = True Then TotalWidth = TotalWidth + GetSystemMetrics(SM_CXVSCROLL)
        If TotalWidth < (UserControl.ScaleWidth - 1) Then
            Dim CountSpring As Long
            For i = 1 To PropShadowPanelsCount
                If PropShadowPanels(i).AutoSize = SbrPanelAutoSizeSpring And PropShadowPanels(i).Visible = True Then CountSpring = CountSpring + 1
            Next i
            If CountSpring > 0 Then
                Dim WidthPerSpring As Long, Remainder As Long
                WidthPerSpring = ((UserControl.ScaleWidth - 1) - TotalWidth) / CountSpring
                Remainder = ((UserControl.ScaleWidth - 1) - TotalWidth) - (WidthPerSpring * CountSpring)
                For i = PropShadowPanelsCount To 1 Step -1
                    If PropShadowPanels(i).AutoSize = SbrPanelAutoSizeSpring And PropShadowPanels(i).Visible = True Then
                        Parts(i - 1) = Parts(i - 1) + WidthPerSpring
                        If Remainder <> 0 Then
                            Parts(i - 1) = Parts(i - 1) + Remainder
                            Remainder = 0
                        End If
                    End If
                Next i
            End If
        End If
        TotalWidth = Borders(SBB_HORIZONTAL)
        For i = 1 To PropShadowPanelsCount
            With PropShadowPanels(i)
            .Left = TotalWidth
            .ActualWidth = Parts(i - 1)
            TotalWidth = TotalWidth + Parts(i - 1) + Borders(SBB_DIVIDER)
            Parts(i - 1) = .Left + .ActualWidth
            End With
        Next i
        SendMessage StatusBarHandle, SB_SETPARTS, PropShadowPanelsCount, ByVal VarPtr(Parts(0))
    Else
        ReDim Parts(0) As Long
        Parts(0) = -1
        SendMessage StatusBarHandle, SB_SETPARTS, 1, ByVal VarPtr(Parts(0))
        SendMessage StatusBarHandle, SB_SETTEXT, 0, ByVal 0&
    End If
End If
End Sub

Private Function GetGoodWidth(ByVal Index As Long) As Long
If StatusBarHandle <> 0 Then
    GetGoodWidth = PropShadowPanels(Index).MinWidth
    If PropShadowPanels(Index).Visible = True Then
        Select Case PropShadowPanels(Index).AutoSize
            Case SbrPanelAutoSizeNone
                If PropShadowPanels(Index).FixedWidth > -1 Then GetGoodWidth = PropShadowPanels(Index).FixedWidth
            Case SbrPanelAutoSizeContent
                Dim Width As Long
                Width = GetTextWidth(Index)
                If Width > GetGoodWidth Then GetGoodWidth = Width
                If Not PropShadowPanels(Index).Picture Is Nothing Then
                    If PropShadowPanels(Index).Picture.Handle <> 0 Then GetGoodWidth = GetGoodWidth + CHimetricToPixel_X(PropShadowPanels(Index).Picture.Width) + 2
                End If
        End Select
    Else
        GetGoodWidth = 0
    End If
End If
End Function

Private Function GetTextWidth(ByVal Index As Long) As Long
If StatusBarHandle <> 0 And StatusBarFontHandle <> 0 Then
    Dim hDC As Long
    Dim hFontOld As Long
    Dim Size As SIZEAPI
    hDC = GetDC(StatusBarHandle)
    If hDC <> 0 Then
        If PropShadowPanels(Index).Bold = False Or StatusBarBoldFontHandle = 0 Then
            hFontOld = SelectObject(hDC, StatusBarFontHandle)
        Else
            hFontOld = SelectObject(hDC, StatusBarBoldFontHandle)
        End If
        If hFontOld <> 0 Then
            Dim Text As String
            Text = PropShadowPanels(Index).DisplayText
            GetTextExtentPoint32 hDC, StrPtr(Text), Len(Text), Size
            GetTextWidth = Size.CX + 8
            SelectObject hDC, hFontOld
        End If
        ReleaseDC StatusBarHandle, hDC
    End If
End If
End Function

Private Sub GetPanelRect(ByVal Index As Long, ByRef RC As RECT)
If StatusBarHandle <> 0 Then
    SendMessage StatusBarHandle, SB_GETRECT, Index - 1, ByVal VarPtr(RC)
    If ComCtlsSupportLevel() = 1 Then ' Bugfix for Windows XP
        If Me.IncludesSizeGrip = True Then
            Dim Parts() As Long
            ReDim Parts(0 To (PropShadowPanelsCount - 1)) As Long
            SendMessage StatusBarHandle, SB_GETPARTS, PropShadowPanelsCount, ByVal VarPtr(Parts(0))
            RC.Right = Parts(Index - 1)
        End If
    End If
End If
End Sub

Private Sub SetPanels()
If StatusBarHandle <> 0 And PropShadowPanelsCount > 0 Then
    Dim i As Long
    For i = 1 To UBound(PropShadowPanels())
        Call SetPanelText(i)
        If PropShowTips = True Then Call SetPanelToolTipText(i)
    Next i
End If
End Sub

Private Sub SetPanelText(ByVal Index As Long)
If StatusBarHandle <> 0 Then
    Dim BevelStyle As Long
    Select Case PropShadowPanels(Index).Bevel
        Case SbrPanelBevelFlat
            BevelStyle = SBT_NOBORDERS
        Case SbrPanelBevelInset
            BevelStyle = 0
        Case SbrPanelBevelRaised
            BevelStyle = SBT_POPOUT
    End Select
    SendMessage StatusBarHandle, SB_SETTEXT, Index - 1 Or BevelStyle Or SBT_OWNERDRAW, ByVal 0&
End If
End Sub

Private Sub SetPanelToolTipText(ByVal Index As Long)
If StatusBarHandle <> 0 And StatusBarToolTipHandle <> 0 Then
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = StatusBarHandle
    .uId = PropShadowPanels(Index).ToolTipID
    If SendMessage(StatusBarToolTipHandle, TTM_GETTOOLINFO, 0, ByVal VarPtr(TI)) <> 0 Then
        .uFlags = TTF_SUBCLASS Or TTF_PARSELINKS
        If PropRightToLeft = True And PropRightToLeftLayout = False Then .uFlags = .uFlags Or TTF_RTLREADING
        .lpszText = StrPtr(PropShadowPanels(Index).ToolTipText)
        Call GetPanelRect(Index, .RC)
        SendMessage StatusBarToolTipHandle, TTM_SETTOOLINFO, 0, ByVal VarPtr(TI)
        SendMessage StatusBarToolTipHandle, TTM_UPDATE, 0, ByVal 0&
    Else
        .uFlags = TTF_SUBCLASS Or TTF_PARSELINKS
        If PropRightToLeft = True And PropRightToLeftLayout = False Then .uFlags = .uFlags Or TTF_RTLREADING
        .lpszText = StrPtr(PropShadowPanels(Index).ToolTipText)
        Call GetPanelRect(Index, .RC)
        SendMessage StatusBarToolTipHandle, TTM_ADDTOOL, 0, ByVal VarPtr(TI)
    End If
    End With
End If
End Sub

Private Sub UpdateToolTipRects()
If StatusBarHandle <> 0 And StatusBarToolTipHandle <> 0 And PropShadowPanelsCount > 0 Then
    Dim TI As TOOLINFO
    With TI
    .cbSize = LenB(TI)
    .hWnd = StatusBarHandle
    Dim i As Long
    For i = 1 To UBound(PropShadowPanels())
        .uId = PropShadowPanels(i).ToolTipID
        Call GetPanelRect(i, .RC)
        SendMessage StatusBarToolTipHandle, TTM_NEWTOOLRECT, 0, ByVal VarPtr(TI)
    Next i
    End With
End If
End Sub

Private Sub CheckTimer()
If StatusBarHandle <> 0 And PropShadowPanelsCount > 0 Then
    TimerUpdatePanels.Enabled = Not StatusBarDesignMode
Else
    TimerUpdatePanels.Enabled = False
End If
End Sub

Private Sub SetVisualStylesToolTip()
If StatusBarHandle <> 0 Then
    If StatusBarToolTipHandle <> 0 And EnabledVisualStyles() = True Then
        If PropVisualStyles = True Then
            ActivateVisualStyles StatusBarToolTipHandle
        Else
            RemoveVisualStyles StatusBarToolTipHandle
        End If
    End If
End If
End Sub

Private Function NextToolTipID() As Long
Static ID As Long
ID = ID + 1
NextToolTipID = ID
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
        If PropDoubleBuffer = True And (StatusBarDoubleBufferEraseBkgDC <> wParam Or StatusBarDoubleBufferEraseBkgDC = 0) And WindowFromDC(wParam) = hWnd Then
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
                    StatusBarDoubleBufferEraseBkgDC = hDCBmp
                    SendMessage hWnd, WM_PRINT, hDCBmp, ByVal PRF_CLIENT Or PRF_ERASEBKGND
                    StatusBarDoubleBufferEraseBkgDC = 0
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
                StatusBarIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                StatusBarIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                StatusBarIsClick = True
            Case WM_MOUSEMOVE
                If StatusBarMouseOver = False And PropMouseTrack = True Then
                    StatusBarMouseOver = True
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
                If StatusBarIsClick = True Then
                    StatusBarIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If StatusBarMouseOver = True Then
            StatusBarMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_SHOWWINDOW
        If StatusBarSizeGripAllowable = True Then
            Dim Align As Integer
            If StatusBarAlignable = True Then Align = Extender.Align Else Align = vbAlignNone
            If Align <> vbAlignBottom Then
                StatusBarSizeGripAllowable = False
                Call ReCreateStatusBar
            End If
        End If
    Case WM_WINDOWPOSCHANGED
        Static PrevWndContainer As Long
        If StatusBarAlignable = True Then
            If PrevWndContainer <> UserControl.ContainerHwnd And PrevWndContainer <> 0 Then
                If Not StatusBarSizeGripAllowable = CBool(((GetWindowLong(UserControl.ContainerHwnd, GWL_STYLE) And WS_THICKFRAME) = WS_THICKFRAME) And Extender.Align = vbAlignBottom) Then
                    StatusBarSizeGripAllowable = Not StatusBarSizeGripAllowable
                    Call ReCreateStatusBar
                End If
            End If
            PrevWndContainer = UserControl.ContainerHwnd
        Else
            If StatusBarSizeGripAllowable = True Then
                StatusBarSizeGripAllowable = False
                Call ReCreateStatusBar
            End If
            PrevWndContainer = 0
        End If
    Case WM_DRAWITEM
        Dim DIS As DRAWITEMSTRUCT
        CopyMemory DIS, ByVal lParam, LenB(DIS)
        If DIS.hWndItem = StatusBarHandle Then
            Call DrawPanel(DIS.ItemID + 1, DIS.hDC, DIS.RCItem)
            WindowProcUserControl = 1
            Exit Function
        End If
    Case WM_NOTIFY
        Dim NM As NMHDR, NMM As NMMOUSE
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = StatusBarHandle Then
            Select Case NM.Code
                Case SBN_SIMPLEMODECHANGE
                    RaiseEvent StyleChange
                Case NM_CLICK, NM_RCLICK
                    If StatusBarIsClick = True Then
                        CopyMemory NMM, ByVal lParam, LenB(NMM)
                        If NMM.dwItemSpec >= 0 Then
                            If NM.Code = NM_CLICK Then
                                RaiseEvent PanelClick(Me.Panels(NMM.dwItemSpec + 1), vbLeftButton)
                            ElseIf NM.Code = NM_RCLICK Then
                                RaiseEvent PanelClick(Me.Panels(NMM.dwItemSpec + 1), vbRightButton)
                            End If
                        End If
                    End If
                Case NM_DBLCLK, NM_RDBLCLK
                    CopyMemory NMM, ByVal lParam, LenB(NMM)
                    If NMM.dwItemSpec >= 0 Then
                        If NM.Code = NM_DBLCLK Then
                            RaiseEvent PanelDblClick(Me.Panels(NMM.dwItemSpec + 1), vbLeftButton)
                        ElseIf NM.Code = NM_RDBLCLK Then
                            RaiseEvent PanelDblClick(Me.Panels(NMM.dwItemSpec + 1), vbRightButton)
                        End If
                    End If
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
If wMsg = WM_DRAWITEM Then
    Dim DIS As DRAWITEMSTRUCT
    CopyMemory DIS, ByVal lParam, LenB(DIS)
    If DIS.hWndItem = StatusBarHandle Then
        Call DrawPanel(DIS.ItemID + 1, DIS.hDC, DIS.RCItem)
        WindowProcUserControlDesignMode = 1
        Exit Function
    End If
End If
WindowProcUserControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
