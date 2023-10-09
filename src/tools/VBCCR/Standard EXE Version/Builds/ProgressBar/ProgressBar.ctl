VERSION 5.00
Begin VB.UserControl ProgressBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "ProgressBar.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "ProgressBar.ctx":004C
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private PrbOrientationHorizontal, PrbOrientationVertical
Private PrbScrollingStandard, PrbScrollingSmooth, PrbScrollingMarquee
Private PrbStateNormal, PrbStateError, PrbStatePaused
#End If
Public Enum PrbOrientationConstants
PrbOrientationHorizontal = 0
PrbOrientationVertical = 1
End Enum
Public Enum PrbScrollingConstants
PrbScrollingStandard = 0
PrbScrollingSmooth = 1
PrbScrollingMarquee = 2
End Enum
Private Const PBST_NORMAL As Long = 1
Private Const PBST_ERROR As Long = 2
Private Const PBST_PAUSED As Long = 3
Public Enum PrbStateConstants
PrbStateNormal = PBST_NORMAL
PrbStateError = PBST_ERROR
PrbStatePaused = PBST_PAUSED
End Enum
Private Const TBPF_NOPROGRESS As Long = 0
Private Const TBPF_INDETERMINATE As Long = 1
Private Const TBPF_NORMAL As Long = 2
Private Const TBPF_ERROR As Long = 4
Private Const TBPF_PAUSED As Long = 8
Private Enum VTableIndexITaskBarList3Constants
' Ignore : ITaskBarList3QueryInterface = 1
' Ignore : ITaskBarList3AddRef = 2
' Ignore : ITaskBarList3Release = 3
VTableIndexITaskBarList3HrInit = 4
' Ignore : ITaskBarList3AddTab = 5
' Ignore : ITaskBarList3DeleteTab = 6
' Ignore : ITaskBarList3ActivateTab = 7
' Ignore : ITaskBarList3SetActiveAlt = 8
' Ignore : ITaskBarList3MarkFullscreenWindow = 9
VTableIndexITaskBarList3SetProgressValue = 10
VTableIndexITaskBarList3SetProgressState = 11
' Ignore : ITaskBarList3RegisterTab = 12
' Ignore : ITaskBarList3UnregisterTab = 13
' Ignore : ITaskBarList3SetTabOrder = 14
' Ignore : ITaskBarList3SetTabActive = 15
' Ignore : ITaskBarList3ThumbBarAddButtons = 16
' Ignore : ITaskBarList3ThumbBarUpdateButtons = 17
' Ignore : ITaskBarList3ThumbBarSetImageList = 18
' Ignore : ITaskBarList3SetOverlayIcon = 19
' Ignore : ITaskBarList3SetThumbnailTooltip = 20
' Ignore : ITaskBarList3SetThumbnailClip = 21
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type PBRANGE
Min As Long
Max As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
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
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As LongPtr, ByRef pCLSID As Any) As Long
Private Declare PtrSafe Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hWnd As LongPtr, ByVal gaFlags As Long) As LongPtr
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetDoubleClickTime Lib "user32" () As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetClassLong Lib "user32" Alias "GetClassLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As LongPtr, ByVal lpchText As LongPtr, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hDC As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nBkMode As Long) As Long
#Else
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
#End If
Private Const ICC_PROGRESS_CLASS As Long = &H20
Private Const CLSID_ITaskBarList As String = "{56FDF344-FD6D-11D0-958A-006097C9A090}"
Private Const IID_ITaskBarList3 As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
Private Const CLSCTX_INPROC_SERVER As Long = 1, S_OK As Long = 0
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const DT_CENTER As Long = &H1
Private Const DT_VCENTER As Long = &H4
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_NOCLIP As Long = &H100
Private Const DT_RTLREADING As Long = &H20000
Private Const GCL_STYLE As Long = (-26)
Private Const CS_DBLCLKS As Long = &H8
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const SW_HIDE As Long = &H0
Private Const GA_ROOT As Long = 2
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
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_SETFONT As Long = &H30
Private Const WM_GETFONT As Long = &H31
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_PAINT As Long = &HF
Private Const WM_PRINTCLIENT As Long = &H318
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETBKCOLOR As Long = (CCM_FIRST + 1)
Private Const WM_USER As Long = &H400
Private Const PBM_SETBKCOLOR As Long = CCM_SETBKCOLOR
Private Const PBM_SETRANGE As Long = (WM_USER + 1) ' 16 bit
Private Const PBM_SETPOS As Long = (WM_USER + 2)
Private Const PBM_DELTAPOS As Long = (WM_USER + 3)
Private Const PBM_SETSTEP As Long = (WM_USER + 4)
Private Const PBM_STEPIT As Long = (WM_USER + 5)
Private Const PBM_SETRANGE32 As Long = (WM_USER + 6)
Private Const PBM_GETRANGE As Long = (WM_USER + 7) ' 16 bit
Private Const PBM_GETPOS As Long = (WM_USER + 8)
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Const PBM_SETMARQUEE As Long = (WM_USER + 10)
Private Const PBM_GETSTEP As Long = (WM_USER + 13)
Private Const PBM_SETSTATE As Long = (WM_USER + 16)
Private Const PBM_GETSTATE As Long = (WM_USER + 17)
Private Const PBS_SMOOTH As Long = &H1
Private Const PBS_VERTICAL As Long = &H4
Private Const PBS_MARQUEE As Long = &H8
Private Const PBS_SMOOTHREVERSE As Long = &H10
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IPerPropertyBrowsingVB
Private ProgressBarHandle As LongPtr
Private ProgressBarFontHandle As LongPtr
Private ProgressBarITaskBarList3 As IUnknown
Private ProgressBarIsClick As Boolean
Private ProgressBarMouseOver As Boolean
Private ProgressBarDesignMode As Boolean
Private ProgressBarDblClickSupported As Boolean, ProgressBarIsDblClick As Boolean
Private ProgressBarDblClickTime As Long, ProgressBarDblClickTickCount As Double
Private ProgressBarDblClickCX As Long, ProgressBarDblClickCY As Long
Private ProgressBarDblClickX As Long, ProgressBarDblClickY As Long
Private ProgressBarAlignable As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropRange As PBRANGE
Private PropValue As Long
Private PropStep As Integer, PropStepAutoReset As Boolean
Private PropMarqueeAnimation As Boolean, PropMarqueeSpeed As Long
Private PropOrientation As PrbOrientationConstants
Private PropScrolling As PrbScrollingConstants
Private PropSmoothReverse As Boolean
Private PropBackColor As OLE_COLOR
Private PropForeColor As OLE_COLOR
Private PropState As PrbStateConstants
Private PropShowInTaskBar As Boolean
Private PropText As String
Private PropTextColor As OLE_COLOR

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
Call ComCtlsInitCC(ICC_PROGRESS_CLASS)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ProgressBarDblClickTime = GetDoubleClickTime()
Const SM_CXDOUBLECLK As Long = 36
Const SM_CYDOUBLECLK As Long = 37
ProgressBarDblClickCX = GetSystemMetrics(SM_CXDOUBLECLK)
ProgressBarDblClickCY = GetSystemMetrics(SM_CYDOUBLECLK)
End Sub

Private Sub UserControl_Show()
Static Done As Boolean
If PropShowInTaskBar = True Then Me.ShowInTaskBar = True
Done = True
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then ProgressBarAlignable = False Else ProgressBarAlignable = True
ProgressBarDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropRange.Min = 0
PropRange.Max = 100
PropValue = 0
PropStep = 10
PropStepAutoReset = True
PropMarqueeAnimation = False
PropMarqueeSpeed = 80
PropOrientation = PrbOrientationHorizontal
PropScrolling = PrbScrollingStandard
PropSmoothReverse = False
PropBackColor = vbButtonFace
PropForeColor = vbHighlight
PropState = PrbStateNormal
PropShowInTaskBar = False
PropText = vbNullString
PropTextColor = vbWindowText
Call CreateProgressBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then ProgressBarAlignable = False Else ProgressBarAlignable = True
ProgressBarDesignMode = Not Ambient.UserMode
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
PropRange.Min = .ReadProperty("Min", 0)
PropRange.Max = .ReadProperty("Max", 100)
PropValue = .ReadProperty("Value", 0)
PropStep = .ReadProperty("Step", 1)
PropStepAutoReset = .ReadProperty("StepAutoReset", True)
PropMarqueeAnimation = .ReadProperty("MarqueeAnimation", False)
PropMarqueeSpeed = .ReadProperty("MarqueeSpeed", 80)
PropOrientation = .ReadProperty("Orientation", PrbOrientationHorizontal)
PropScrolling = .ReadProperty("Scrolling", PrbScrollingStandard)
PropSmoothReverse = .ReadProperty("SmoothReverse", PropSmoothReverse)
PropBackColor = .ReadProperty("BackColor", vbButtonFace)
PropForeColor = .ReadProperty("ForeColor", vbHighlight)
PropState = .ReadProperty("State", PrbStateNormal)
PropShowInTaskBar = .ReadProperty("ShowInTaskBar", False)
PropText = VarToStr(.ReadProperty("Text", vbNullString))
PropTextColor = .ReadProperty("TextColor", vbWindowText)
End With
Call CreateProgressBar
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
.WriteProperty "Min", PropRange.Min, 0
.WriteProperty "Max", PropRange.Max, 100
.WriteProperty "Value", PropValue, 0
.WriteProperty "Step", PropStep, 1
.WriteProperty "StepAutoReset", PropStepAutoReset, True
.WriteProperty "MarqueeAnimation", PropMarqueeAnimation, False
.WriteProperty "MarqueeSpeed", PropMarqueeSpeed, 80
.WriteProperty "Orientation", PropOrientation, PrbOrientationHorizontal
.WriteProperty "Scrolling", PropScrolling, PrbScrollingStandard
.WriteProperty "SmoothReverse", PropSmoothReverse, False
.WriteProperty "BackColor", PropBackColor, vbButtonFace
.WriteProperty "ForeColor", PropForeColor, vbHighlight
.WriteProperty "State", PropState, PrbStateNormal
.WriteProperty "ShowInTaskBar", PropShowInTaskBar, False
.WriteProperty "Text", StrToVar(PropText), vbNullString
.WriteProperty "TextColor", PropTextColor, vbWindowText
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
If ProgressBarAlignable = True Then Align = .Align Else Align = vbAlignNone
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
If ProgressBarHandle <> NULL_PTR Then MoveWindow ProgressBarHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyProgressBar
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
hWnd = ProgressBarHandle
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
OldFontHandle = ProgressBarFontHandle
ProgressBarFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ProgressBarHandle <> NULL_PTR Then
    SendMessage ProgressBarHandle, WM_SETFONT, ProgressBarFontHandle, ByVal 1&
    If Not PropText = vbNullString Then InvalidateRect ProgressBarHandle, ByVal NULL_PTR, 1
End If
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As LongPtr
OldFontHandle = ProgressBarFontHandle
ProgressBarFontHandle = CreateGDIFontFromOLEFont(PropFont)
If ProgressBarHandle <> NULL_PTR Then
    SendMessage ProgressBarHandle, WM_SETFONT, ProgressBarFontHandle, ByVal 1&
    If Not PropText = vbNullString Then InvalidateRect ProgressBarHandle, ByVal NULL_PTR, 1
End If
If OldFontHandle <> NULL_PTR Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If ProgressBarHandle <> NULL_PTR And EnabledVisualStyles() = True Then
    Dim dwExStyle As Long, dwExStyleOld As Long
    dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    dwExStyleOld = dwExStyle
    If PropVisualStyles = True Then
        ActivateVisualStyles ProgressBarHandle
        If (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle And Not WS_EX_STATICEDGE
    Else
        RemoveVisualStyles ProgressBarHandle
        If Not (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle Or WS_EX_STATICEDGE
    End If
    If dwExStyle <> dwExStyleOld Then
        SetWindowLong ProgressBarHandle, GWL_EXSTYLE, dwExStyle
        Call ComCtlsFrameChanged(ProgressBarHandle)
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
If ProgressBarDesignMode = False Then Call RefreshMousePointer
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
        If ProgressBarDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If ProgressBarDesignMode = False Then Call RefreshMousePointer
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
If ProgressBarDesignMode = False Then Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
If ProgressBarHandle <> NULL_PTR Then Call ComCtlsSetRightToLeft(ProgressBarHandle, dwMask)
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

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum position."
If ProgressBarHandle <> NULL_PTR Then
    Min = CLng(SendMessage(ProgressBarHandle, PBM_GETRANGE, 1, ByVal 0&))
Else
    Min = PropRange.Min
End If
End Property

Public Property Let Min(ByVal Value As Long)
If Value < Me.Max Then
    PropRange.Min = Value
    PropRange.Max = Me.Max
    If PropValue < PropRange.Min Then PropValue = PropRange.Min
Else
    If ProgressBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If ProgressBarHandle <> NULL_PTR Then SendMessage ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, ByVal PropRange.Max
UserControl.PropertyChanged "Min"
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum position."
If ProgressBarHandle = NULL_PTR Then
    Max = CLng(SendMessage(ProgressBarHandle, PBM_GETRANGE, 0, ByVal 0&))
Else
    Max = PropRange.Max
End If
End Property

Public Property Let Max(ByVal Value As Long)
If Value > Me.Min Then
    PropRange.Min = Me.Min
    PropRange.Max = Value
    If PropValue > PropRange.Max Then PropValue = PropRange.Max
Else
    If ProgressBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If ProgressBarHandle <> NULL_PTR Then SendMessage ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, ByVal PropRange.Max
UserControl.PropertyChanged "Max"
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the current position."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "123c"
If ProgressBarHandle <> NULL_PTR And (PropScrolling <> PrbScrollingMarquee Or ComCtlsSupportLevel() = 0) Then
    Value = CLng(SendMessage(ProgressBarHandle, PBM_GETPOS, 0, ByVal 0&))
Else
    Value = PropValue
End If
End Property

Public Property Let Value(ByVal NewValue As Long)
If NewValue > Me.Max Then
    NewValue = Me.Max
ElseIf NewValue < Me.Min Then
    NewValue = Me.Min
End If
Dim Changed As Boolean
Changed = CBool(Me.Value <> NewValue)
PropValue = NewValue
If ProgressBarHandle <> NULL_PTR And (PropScrolling <> PrbScrollingMarquee Or ComCtlsSupportLevel() = 0) Then SendMessage ProgressBarHandle, PBM_SETPOS, PropValue, ByVal 0&
UserControl.PropertyChanged "Value"
If Changed = True Then
    On Error Resume Next
    UserControl.Extender.DataChanged = True
    On Error GoTo 0
    Call CheckTaskBarProgress
    RaiseEvent Change
End If
End Property

Public Property Get Step() As Long
Attribute Step.VB_Description = "Returns/sets the step value for the 'StepIt' procedure."
If ProgressBarHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then
    Step = CLng(SendMessage(ProgressBarHandle, PBM_GETSTEP, 0, ByVal 0&))
Else
    Step = PropStep
End If
End Property

Public Property Let Step(ByVal Value As Long)
PropStep = Value
If ProgressBarHandle <> NULL_PTR Then SendMessage ProgressBarHandle, PBM_SETSTEP, PropStep, ByVal 0&
UserControl.PropertyChanged "Step"
End Property

Public Property Get StepAutoReset() As Boolean
Attribute StepAutoReset.VB_Description = "Returns/sets a value that determines whether the position will be automatically reset when the maximum is exceeded or not. Only applicable for the 'StepIt' procedure."
StepAutoReset = PropStepAutoReset
End Property

Public Property Let StepAutoReset(ByVal Value As Boolean)
PropStepAutoReset = Value
UserControl.PropertyChanged "StepAutoReset"
End Property

Public Property Get MarqueeAnimation() As Boolean
Attribute MarqueeAnimation.VB_Description = "Returns/sets a value that determines whether the marquee animation is on or off. Requires comctl32.dll version 6.0 or higher."
MarqueeAnimation = PropMarqueeAnimation
End Property

Public Property Let MarqueeAnimation(ByVal Value As Boolean)
PropMarqueeAnimation = Value
If ProgressBarHandle <> NULL_PTR And ComCtlsSupportLevel() >= 1 Then SendMessage ProgressBarHandle, PBM_SETMARQUEE, IIf(PropMarqueeAnimation = True, 1, 0), ByVal PropMarqueeSpeed
Call CheckTaskBarProgress
UserControl.PropertyChanged "MarqueeAnimation"
End Property

Public Property Get MarqueeSpeed() As Long
Attribute MarqueeSpeed.VB_Description = "Returns/sets the speed of the marquee animation. That means the time, in milliseconds, between marquee animation updates. Requires comctl32.dll version 6.0 or higher."
MarqueeSpeed = PropMarqueeSpeed
End Property

Public Property Let MarqueeSpeed(ByVal Value As Long)
If Value > 0 Then
    PropMarqueeSpeed = Value
Else
    If ProgressBarDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If ProgressBarHandle <> NULL_PTR And ComCtlsSupportLevel() >= 1 Then SendMessage ProgressBarHandle, PBM_SETMARQUEE, IIf(PropMarqueeAnimation = True, 1, 0), ByVal PropMarqueeSpeed
UserControl.PropertyChanged "MarqueeSpeed"
End Property

Public Property Get Orientation() As PrbOrientationConstants
Attribute Orientation.VB_Description = "Returns/sets the orientation."
Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As PrbOrientationConstants)
Select Case Value
    Case PrbOrientationHorizontal, PrbOrientationVertical
        With UserControl
        Dim Align As Integer
        If ProgressBarAlignable = True Then Align = .Extender.Align Else Align = vbAlignNone
        If Align = vbAlignNone And PropOrientation <> Value Then
            .Extender.Move .Extender.Left, .Extender.Top, .Extender.Height, .Extender.Width
        End If
        End With
        PropOrientation = Value
    Case Else
        Err.Raise 380
End Select
If ProgressBarHandle <> NULL_PTR Then Call ReCreateProgressBar
UserControl.PropertyChanged "Orientation"
End Property

Public Property Get Scrolling() As PrbScrollingConstants
Attribute Scrolling.VB_Description = "Returns/sets the scrolling."
Scrolling = PropScrolling
End Property

Public Property Let Scrolling(ByVal Value As PrbScrollingConstants)
Select Case Value
    Case PrbScrollingStandard, PrbScrollingSmooth, PrbScrollingMarquee
        PropScrolling = Value
    Case Else
        Err.Raise 380
End Select
If ProgressBarHandle <> NULL_PTR Then Call ReCreateProgressBar
Call CheckTaskBarProgress
UserControl.PropertyChanged "Scrolling"
End Property

Public Property Get SmoothReverse() As Boolean
Attribute SmoothReverse.VB_Description = "Returns/sets a value that determines the animation behavior when moving backward. If this is set, then a smooth transition will occur, otherwise it will jump to the lower value. Requires comctl32.dll version 6.1 or higher."
SmoothReverse = PropSmoothReverse
End Property

Public Property Let SmoothReverse(ByVal Value As Boolean)
PropSmoothReverse = Value
If ProgressBarHandle <> NULL_PTR And ComCtlsSupportLevel() >= 1 Then Call ReCreateProgressBar
UserControl.PropertyChanged "SmoothReverse"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object. This property is ignored if the version of comctl32.dll is 6.0 or higher and the visual styles property is set to true."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If ProgressBarHandle <> NULL_PTR Then SendMessage ProgressBarHandle, PBM_SETBKCOLOR, 0, ByVal WinColor(PropBackColor)
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object. This property is ignored if the version of comctl32.dll is 6.0 or higher and the visual styles property is set to true."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
If ProgressBarHandle <> NULL_PTR Then SendMessage ProgressBarHandle, PBM_SETBARCOLOR, 0, ByVal WinColor(PropForeColor)
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get State() As PrbStateConstants
Attribute State.VB_Description = "Returns/sets the state of the progress bar. Requires comctl32.dll version 6.1 or higher."
If ProgressBarHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then
    State = CLng(SendMessage(ProgressBarHandle, PBM_GETSTATE, 0, ByVal 0&))
Else
    State = PropState
End If
End Property

Public Property Let State(ByVal Value As PrbStateConstants)
Select Case Value
    Case PrbStateNormal, PrbStateError, PrbStatePaused
        PropState = Value
    Case Else
        Err.Raise 380
End Select
If ProgressBarHandle <> NULL_PTR And ComCtlsSupportLevel() >= 2 Then SendMessage ProgressBarHandle, PBM_SETSTATE, PropState, ByVal 0&
Call CheckTaskBarProgress
UserControl.PropertyChanged "State"
End Property

Public Property Get ShowInTaskBar() As Boolean
Attribute ShowInTaskBar.VB_Description = "Returns/sets a value that indicates if the progress state and value appears in the Windows 95 taskbar. Requires comctl32.dll version 6.1 or higher."
ShowInTaskBar = PropShowInTaskBar
End Property

Public Property Let ShowInTaskBar(ByVal Value As Boolean)
PropShowInTaskBar = Value
If ProgressBarDesignMode = False And ComCtlsSupportLevel() >= 2 Then
    If ProgressBarITaskBarList3 Is Nothing Then Set ProgressBarITaskBarList3 = CreateITaskBarList3()
    If PropShowInTaskBar = True Then
        Call CheckTaskBarProgress
    Else
        If ProgressBarHandle <> NULL_PTR Then
            If Not ProgressBarITaskBarList3 Is Nothing Then
                Dim hWnd As LongPtr
                hWnd = GetAncestor(ProgressBarHandle, GA_ROOT)
                If hWnd <> NULL_PTR Then VTableCall vbEmpty, ObjPtr(ProgressBarITaskBarList3), VTableIndexITaskBarList3SetProgressState, hWnd, TBPF_NOPROGRESS
            End If
        End If
    End If
End If
UserControl.PropertyChanged "ShowInTaskBar"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object. Placeholders: {0} = Value, {1} = Min, {2} = Max and {3} = Percent value between 0 and 100."
Text = PropText
End Property

Public Property Let Text(ByVal Value As String)
If PropText = Value Then Exit Property
If ProgressBarDesignMode = True Then
    If ProgressBarHandle <> NULL_PTR Then
        If Value = vbNullString And Not PropText = vbNullString Then
            Call ComCtlsRemoveSubclass(ProgressBarHandle)
        ElseIf Not Value = vbNullString And PropText = vbNullString Then
            Call ComCtlsSetSubclass(ProgressBarHandle, Me, 2)
        End If
    End If
End If
PropText = Value
If ProgressBarHandle <> NULL_PTR Then InvalidateRect ProgressBarHandle, ByVal NULL_PTR, 1
UserControl.PropertyChanged "Text"
End Property

Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
TextColor = PropTextColor
End Property

Public Property Let TextColor(ByVal Value As OLE_COLOR)
PropTextColor = Value
If ProgressBarHandle <> NULL_PTR Then InvalidateRect ProgressBarHandle, ByVal NULL_PTR, 1
UserControl.PropertyChanged "TextColor"
End Property

Private Sub CreateProgressBar()
If ProgressBarHandle <> NULL_PTR Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If PropOrientation = PrbOrientationVertical Then dwStyle = dwStyle Or PBS_VERTICAL
Select Case PropScrolling
    Case PrbScrollingSmooth
        dwStyle = dwStyle Or PBS_SMOOTH
    Case PrbScrollingMarquee
        If ComCtlsSupportLevel() >= 1 Then
            dwStyle = dwStyle Or PBS_MARQUEE
        Else
            If ProgressBarDesignMode = False Then PropScrolling = PrbScrollingStandard
        End If
End Select
If PropSmoothReverse = True Then If ComCtlsSupportLevel() >= 1 Then dwStyle = dwStyle Or PBS_SMOOTHREVERSE
ProgressBarHandle = CreateWindowEx(dwExStyle, StrPtr("msctls_progress32"), NULL_PTR, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
If ProgressBarHandle <> NULL_PTR Then SendMessage ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, ByVal PropRange.Max
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Value = PropValue
Me.Step = PropStep
Me.MarqueeAnimation = PropMarqueeAnimation
Me.BackColor = PropBackColor
Me.ForeColor = PropForeColor
Me.State = PropState
If ProgressBarDesignMode = False Then
    If ProgressBarHandle <> NULL_PTR Then
        ProgressBarDblClickSupported = CBool((GetClassLong(ProgressBarHandle, GCL_STYLE) And CS_DBLCLKS) <> 0)
        Call ComCtlsSetSubclass(ProgressBarHandle, Me, 1)
    End If
ElseIf Not PropText = vbNullString Then
    If ProgressBarHandle <> NULL_PTR Then Call ComCtlsSetSubclass(ProgressBarHandle, Me, 2)
End If
End Sub

Private Sub ReCreateProgressBar()
If ProgressBarDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Call DestroyProgressBar
    Call CreateProgressBar
    Call UserControl_Resize
    If Locked = True Then LockWindowUpdate NULL_PTR
    Me.Refresh
Else
    Call DestroyProgressBar
    Call CreateProgressBar
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyProgressBar()
If ProgressBarHandle = NULL_PTR Then Exit Sub
Call ComCtlsRemoveSubclass(ProgressBarHandle)
ShowWindow ProgressBarHandle, SW_HIDE
SetParent ProgressBarHandle, NULL_PTR
DestroyWindow ProgressBarHandle
ProgressBarHandle = NULL_PTR
If ProgressBarFontHandle <> NULL_PTR Then
    DeleteObject ProgressBarFontHandle
    ProgressBarFontHandle = NULL_PTR
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub StepIt()
Attribute StepIt.VB_Description = "Advances the current position by the step increment."
If ProgressBarHandle <> NULL_PTR And (PropScrolling <> PrbScrollingMarquee Or ComCtlsSupportLevel() = 0) Then
    Dim Changed As Boolean
    If PropStepAutoReset = True Then
        SendMessage ProgressBarHandle, PBM_STEPIT, 0, ByVal 0&
        Changed = CBool(Me.Value <> PropValue)
    Else
        If (Me.Value + Me.Step) <= Me.Max Then
            SendMessage ProgressBarHandle, PBM_STEPIT, 0, ByVal 0&
            Changed = CBool(Me.Value <> PropValue)
        Else
            Me.Value = Me.Max
        End If
    End If
    If Changed = True Then
        PropValue = Me.Value
        UserControl.PropertyChanged "Value"
        On Error Resume Next
        UserControl.Extender.DataChanged = True
        On Error GoTo 0
        Call CheckTaskBarProgress
        RaiseEvent Change
    End If
End If
End Sub

Public Sub Increment(ByVal Delta As Long)
Attribute Increment.VB_Description = "Advances the current position by a specified increment."
If ProgressBarHandle <> NULL_PTR And (PropScrolling <> PrbScrollingMarquee Or ComCtlsSupportLevel() = 0) Then
    SendMessage ProgressBarHandle, PBM_DELTAPOS, Delta, ByVal 0&
    If Me.Value <> PropValue Then
        PropValue = Me.Value
        UserControl.PropertyChanged "Value"
        On Error Resume Next
        UserControl.Extender.DataChanged = True
        On Error GoTo 0
        Call CheckTaskBarProgress
        RaiseEvent Change
    End If
End If
End Sub

Private Sub CheckTaskBarProgress()
If PropShowInTaskBar = False Or ProgressBarITaskBarList3 Is Nothing Then Exit Sub
If ProgressBarHandle <> NULL_PTR Then
    Dim hWnd As LongPtr
    hWnd = GetAncestor(ProgressBarHandle, GA_ROOT)
    If hWnd <> NULL_PTR Then
        Dim TaskBarState As Long
        If PropScrolling <> PrbScrollingMarquee Then
            Select Case PropState
                Case PrbStateNormal
                    TaskBarState = TBPF_NORMAL
                Case PrbStateError
                    TaskBarState = TBPF_ERROR
                Case PrbStatePaused
                    TaskBarState = TBPF_PAUSED
            End Select
            #If Win64 Then
            VTableCall vbEmpty, ObjPtr(ProgressBarITaskBarList3), VTableIndexITaskBarList3SetProgressValue, hWnd, CLngLng(PropValue), CLngLng(Me.Max - Me.Min)
            #Else
            VTableCall vbEmpty, ObjPtr(ProgressBarITaskBarList3), VTableIndexITaskBarList3SetProgressValue, hWnd, PropValue, 0&, CLng(Me.Max - Me.Min), 0&
            #End If
        Else
            If PropMarqueeAnimation = True Then
                TaskBarState = TBPF_INDETERMINATE
            Else
                TaskBarState = TBPF_NOPROGRESS
            End If
        End If
        VTableCall vbEmpty, ObjPtr(ProgressBarITaskBarList3), VTableIndexITaskBarList3SetProgressState, hWnd, TaskBarState
    End If
End If
End Sub

Private Function CreateITaskBarList3() As IUnknown
Dim CLSID As OLEGuids.OLECLSID, IID As OLEGuids.OLECLSID
On Error Resume Next
CLSIDFromString StrPtr(CLSID_ITaskBarList), CLSID
CLSIDFromString StrPtr(IID_ITaskBarList3), IID
CoCreateInstance CLSID, NULL_PTR, CLSCTX_INPROC_SERVER, IID, CreateITaskBarList3
If Not CreateITaskBarList3 Is Nothing Then
    VTableCall vbEmpty, ObjPtr(CreateITaskBarList3), VTableIndexITaskBarList3HrInit
    If Err.LastDllError <> S_OK Then Set CreateITaskBarList3 = Nothing
End If
End Function

Private Sub TextDraw(ByVal hWnd As LongPtr, ByVal hDC As LongPtr)
If hWnd = NULL_PTR Or hDC = NULL_PTR Then Exit Sub
If PropText = vbNullString Then Exit Sub
Dim hFont As LongPtr, hFontOld As LongPtr
hFont = SendMessage(hWnd, WM_GETFONT, 0, ByVal 0&)
hFontOld = SelectObject(hDC, hFont)
Dim OldBkMode As Long, OldTextColor As Long
OldBkMode = SetBkMode(hDC, 1)
OldTextColor = SetTextColor(hDC, WinColor(PropTextColor))
Dim DrawFlags As Long, RC As RECT
DrawFlags = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
If PropRightToLeft = True And PropRightToLeftLayout = False Then DrawFlags = DrawFlags Or DT_RTLREADING
GetClientRect hWnd, RC
Dim Text As String, Pos As Long
For Pos = 1 To Len(PropText)
    Select Case Mid$(PropText, Pos, 3)
        Case "{0}" ' Value
            Text = Text & CStr(Me.Value)
            Pos = Pos + 2
        Case "{1}" ' Min
            Text = Text & CStr(Me.Min)
            Pos = Pos + 2
        Case "{2}" ' Max
            Text = Text & CStr(Me.Max)
            Pos = Pos + 2
        Case "{3}" ' Percent value between 0 and 100
            Text = Text & Format$((CDbl(Me.Value - Me.Min) / CDbl(Me.Max - Me.Min)) * 100, "0")
            Pos = Pos + 2
        Case Else
            Text = Text & Mid$(PropText, Pos, 1)
    End Select
Next Pos
DrawText hDC, StrPtr(Text), -1, RC, DrawFlags
SetBkMode hDC, OldBkMode
SetTextColor hDC, OldTextColor
If hFontOld <> NULL_PTR Then SelectObject hDC, hFontOld
End Sub

Private Function PtInRect(ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
' Avoid API declare since x64 calling convention aligns 8 bytes per argument.
' So the handling of a ByVal PT being split into two 4-byte arguments will crash.
PtInRect = 0
If X >= lpRect.Left And X < lpRect.Right And Y >= lpRect.Top And Y < lpRect.Bottom Then PtInRect = 1
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
        ISubclass_Message = WindowProcControlDesignMode(hWnd, wMsg, wParam, lParam)
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
    Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
        If ProgressBarDblClickSupported = False Then
            If ProgressBarDblClickTickCount = 0 Then
                ProgressBarDblClickTickCount = CLngToULng(GetTickCount())
                ProgressBarDblClickX = Get_X_lParam(lParam)
                ProgressBarDblClickY = Get_Y_lParam(lParam)
            Else
                If (CLngToULng(GetTickCount()) - ProgressBarDblClickTickCount) <= ProgressBarDblClickTime Then
                    Dim DblClickRect As RECT
                    With DblClickRect
                    .Left = ProgressBarDblClickX - (ProgressBarDblClickCX / 2)
                    .Right = ProgressBarDblClickX + (ProgressBarDblClickCX / 2)
                    .Top = ProgressBarDblClickY - (ProgressBarDblClickCY / 2)
                    .Bottom = ProgressBarDblClickY + (ProgressBarDblClickCY / 2)
                    End With
                    If PtInRect(DblClickRect, Get_X_lParam(lParam), Get_Y_lParam(lParam)) <> 0 Then ProgressBarIsDblClick = True
                End If
                ProgressBarDblClickTickCount = CLngToULng(GetTickCount())
                ProgressBarDblClickX = Get_X_lParam(lParam)
                ProgressBarDblClickY = Get_Y_lParam(lParam)
            End If
            If ProgressBarIsDblClick = True Then
                Select Case wMsg
                    Case WM_LBUTTONDOWN
                        wMsg = WM_LBUTTONDBLCLK
                    Case WM_MBUTTONDOWN
                        wMsg = WM_MBUTTONDBLCLK
                    Case WM_RBUTTONDOWN
                        wMsg = WM_RBUTTONDBLCLK
                End Select
                ProgressBarIsDblClick = False
            End If
        End If
End Select
WindowProcControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_PAINT, WM_PRINTCLIENT
        If Not PropText = vbNullString Then
            If wMsg = WM_PAINT Then
                Dim hDC As LongPtr
                hDC = GetDC(hWnd)
                If hDC <> NULL_PTR Then
                    Call TextDraw(hWnd, hDC)
                    ReleaseDC hWnd, hDC
                End If
            ElseIf wMsg = WM_PRINTCLIENT Then
                If wParam <> 0 Then Call TextDraw(hWnd, wParam)
            End If
        End If
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
                ProgressBarIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                ProgressBarIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                ProgressBarIsClick = True
            Case WM_MOUSEMOVE
                If ProgressBarMouseOver = False And PropMouseTrack = True Then
                    ProgressBarMouseOver = True
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
                If ProgressBarIsClick = True Then
                    ProgressBarIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If ProgressBarMouseOver = True Then
            ProgressBarMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcControlDesignMode(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_PAINT, WM_PRINTCLIENT
        WindowProcControlDesignMode = WindowProcControl(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
WindowProcControlDesignMode = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
Select Case wMsg
    Case WM_DESTROY, WM_NCDESTROY
        Call ComCtlsRemoveSubclass(hWnd)
End Select
End Function
