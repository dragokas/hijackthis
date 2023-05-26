VERSION 5.00
Begin VB.UserControl Pager 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "Pager.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "Pager.ctx":0037
   Begin VB.Timer TimerBuddyControl 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Pager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private PgrOrientationVertical, PgrOrientationHorizontal
Private PgrDirectionUp, PgrDirectionDown, PgrDirectionLeft, PgrDirectionRight
Private PgrButtonTopOrLeft, PgrButtonBottomOrRight
Private PgrButtonStateInvisible, PgrButtonStateNormal, PgrButtonStateGrayed, PgrButtonStateDepressed, PgrButtonStateHot
#End If
Public Enum PgrOrientationConstants
PgrOrientationVertical = 0
PgrOrientationHorizontal = 1
End Enum
Private Const PGF_SCROLLUP As Long = 1
Private Const PGF_SCROLLDOWN As Long = 2
Private Const PGF_SCROLLLEFT As Long = 4
Private Const PGF_SCROLLRIGHT As Long = 8
Public Enum PgrDirectionConstants
PgrDirectionUp = PGF_SCROLLUP
PgrDirectionDown = PGF_SCROLLDOWN
PgrDirectionLeft = PGF_SCROLLLEFT
PgrDirectionRight = PGF_SCROLLRIGHT
End Enum
Private Const PGB_TOPORLEFT As Long = 0
Private Const PGB_BOTTOMORRIGHT As Long = 1
Public Enum PgrButtonConstants
PgrButtonTopOrLeft = PGB_TOPORLEFT
PgrButtonBottomOrRight = PGB_BOTTOMORRIGHT
End Enum
Private Const PGF_INVISIBLE As Long = 0
Private Const PGF_NORMAL As Long = 1
Private Const PGF_GRAYED As Long = 2
Private Const PGF_DEPRESSED As Long = 4
Private Const PGF_HOT As Long = 8
Public Enum PgrButtonStateConstants
PgrButtonStateInvisible = PGF_INVISIBLE
PgrButtonStateNormal = PGF_NORMAL
PgrButtonStateGrayed = PGF_GRAYED
PgrButtonStateDepressed = PGF_DEPRESSED
PgrButtonStateHot = PGF_HOT
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type NMHDR
hWndFrom As Long
IDFrom As Long
Code As Long
End Type
Private Type NMPGCALCSIZE
hdr As NMHDR
dwFlag As Long
iWidth As Long
iHeight As Long
End Type
Private Type NMPGHOTITEM
hdr As NMHDR
IDOld As Long
IDNew As Long
dwFlags As Long
End Type
Private Type NMPGSCROLL
hdr As NMHDR
fwKeys As Integer
RCParent(0 To 15) As Byte
iDir(0 To 1) As Integer
iXPos(0 To 1) As Integer
iYPos(0 To 1) As Integer
iScroll(0 To 1) As Integer
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event Scroll(ByVal Shift As Integer, ByVal Direction As PgrDirectionConstants, ByVal X As Single, ByVal Y As Single, ByRef Delta As Single, ByVal ClientLeft As Single, ByVal ClientTop As Single, ByVal ClientRight As Single, ByVal ClientBottom As Single)
Attribute Scroll.VB_Description = "Occurs whenever the pager control scrolls the buddy control."
Public Event CalcSize(ByRef Width As Single, ByRef Height As Single)
Attribute CalcSize.VB_Description = "Occurs whenever the pager control needs to know the size of the buddy control. Default values are set but you can change them if necessary. You do not need to handle this event in most circumstances."
Public Event HotChanged()
Attribute HotChanged.VB_Description = "Occurrs when the pager control's hot state changes. Requires comctl32.dll version 6.0 or higher."
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
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const ICC_PAGESCROLLER_CLASS As Long = &H1000
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const SW_HIDE As Long = &H0
Private Const GA_PARENT As Long = 1
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const PGM_FIRST As Long = &H1400
Private Const PGM_SETCHILD As Long = (PGM_FIRST + 1)
Private Const PGM_RECALCSIZE As Long = (PGM_FIRST + 2)
Private Const PGM_FORWARDMOUSE As Long = (PGM_FIRST + 3)
Private Const PGM_SETBKCOLOR As Long = (PGM_FIRST + 4)
Private Const PGM_GETBKCOLOR As Long = (PGM_FIRST + 5)
Private Const PGM_SETBORDER As Long = (PGM_FIRST + 6)
Private Const PGM_GETBORDER As Long = (PGM_FIRST + 7)
Private Const PGM_SETPOS As Long = (PGM_FIRST + 8)
Private Const PGM_GETPOS As Long = (PGM_FIRST + 9)
Private Const PGM_SETBUTTONSIZE As Long = (PGM_FIRST + 10)
Private Const PGM_GETBUTTONSIZE As Long = (PGM_FIRST + 11)
Private Const PGM_GETBUTTONSTATE As Long = (PGM_FIRST + 12)
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_GETDROPTARGET As Long = (CCM_FIRST + 4)
Private Const PGM_GETDROPTARGET As Long = CCM_GETDROPTARGET
Private Const PGN_FIRST As Long = (-900)
Private Const PGN_SCROLL As Long = (PGN_FIRST - 1)
Private Const PGN_CALCSIZE As Long = (PGN_FIRST - 2)
Private Const PGN_HOTITEMCHANGE As Long = (PGN_FIRST - 3)
Private Const PGS_VERT As Long = &H0
Private Const PGS_HORZ As Long = &H1
Private Const PGS_AUTOSCROLL As Long = &H2
Private Const PGS_DRAGNDROP As Long = &H4
Private Const HICF_ENTERING As Long = &H10
Private Const HICF_LEAVING As Long = &H20
Private Const PGF_CALCWIDTH As Long = 1
Private Const PGF_CALCHEIGHT As Long = 2
Private Const PGK_SHIFT As Long = 1
Private Const PGK_CONTROL As Long = 2
Private Const PGK_MENU As Long = 4
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IPerPropertyBrowsingVB
Private PagerHandle As Long
Private PagerIsClick As Boolean
Private PagerMouseOver As Boolean
Private PagerDesignMode As Boolean
Private PagerHotItemChangePrevFlags As Long
Private PagerAlignable As Boolean
Private PagerBuddyControlHandle As Long, PagerBuddyControlPrevParent As Long
Private PagerBuddyObjectPointer As Long
Private DispIDMousePointer As Long
Private DispIDBuddyControl As Long, BuddyControlArray() As String
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBuddyName As String, PropBuddyControlInit As Boolean
Private PropBackColor As OLE_COLOR
Private PropOLEDragDropScroll As Boolean
Private PropOrientation As PgrOrientationConstants
Private PropBorderWidth As Long
Private PropAutoScroll As Boolean
Private PropButtonSize As Long

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
ElseIf DispID = DispIDBuddyControl Then
    DisplayName = PropBuddyName
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDMousePointer Then
    Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
    Handled = True
ElseIf DispID = DispIDBuddyControl Then
    On Error GoTo CATCH_EXCEPTION
    Dim ControlEnum As Object, PropUBound As Long, Handle As Long
    PropUBound = UBound(StringsOut())
    ReDim Preserve StringsOut(PropUBound + 1) As String
    ReDim Preserve CookiesOut(PropUBound + 1) As Long
    StringsOut(PropUBound) = "(None)"
    CookiesOut(PropUBound) = PropUBound
    For Each ControlEnum In UserControl.ParentControls
        If ControlIsValid(ControlEnum, Handle) = True Then
            If Handle <> 0 Then
                PropUBound = UBound(StringsOut())
                ReDim Preserve StringsOut(PropUBound + 1) As String
                ReDim Preserve CookiesOut(PropUBound + 1) As Long
                StringsOut(PropUBound) = ProperControlName(ControlEnum)
                CookiesOut(PropUBound) = PropUBound
            End If
        End If
    Next ControlEnum
    PropUBound = UBound(StringsOut())
    ReDim BuddyControlArray(0 To PropUBound) As String
    Dim i As Long
    For i = 0 To PropUBound
        BuddyControlArray(i) = StringsOut(i)
    Next i
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
ElseIf DispID = DispIDBuddyControl Then
    If Cookie < UBound(BuddyControlArray()) Then Value = BuddyControlArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_PAGESCROLLER_CLASS)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim BuddyControlArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDBuddyControl = 0 Then DispIDBuddyControl = GetDispID(Me, "BuddyControl")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then PagerAlignable = False Else PagerAlignable = True
PagerDesignMode = Not Ambient.UserMode
On Error GoTo 0
PropBackColor = vbButtonFace
PropOLEDragDropScroll = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBuddyName = "(None)"
PropOrientation = PgrOrientationVertical
PropBorderWidth = 0
PropAutoScroll = False
PropButtonSize = -1
Call CreatePager
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDBuddyControl = 0 Then DispIDBuddyControl = GetDispID(Me, "BuddyControl")
On Error Resume Next
If UserControl.ParentControls.Count = 0 Then PagerAlignable = False Else PagerAlignable = True
PagerDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
PropBackColor = .ReadProperty("BackColor", vbButtonFace)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropOLEDragDropScroll = .ReadProperty("OLEDragDropScroll", True)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropBuddyName = .ReadProperty("BuddyControl", "(None)")
PropOrientation = .ReadProperty("Orientation", PgrOrientationVertical)
PropBorderWidth = (.ReadProperty("BorderWidth", 0) * PixelsPerDIP_X())
PropAutoScroll = .ReadProperty("AutoScroll", False)
PropButtonSize = .ReadProperty("ButtonSize", -1)
End With
Call CreatePager
If Not PropBuddyName = "(None)" Then TimerBuddyControl.Enabled = Ambient.UserMode
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "BackColor", PropBackColor, vbButtonFace
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "OLEDragDropScroll", PropOLEDragDropScroll, True
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BuddyControl", PropBuddyName, "(None)"
.WriteProperty "Orientation", PropOrientation, PgrOrientationVertical
.WriteProperty "BorderWidth", (PropBorderWidth / PixelsPerDIP_X()), 0
.WriteProperty "AutoScroll", PropAutoScroll, False
.WriteProperty "ButtonSize", PropButtonSize, -1
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
If PagerAlignable = True Then Align = .Align Else Align = vbAlignNone
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
If PagerHandle <> 0 Then MoveWindow PagerHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyPager
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerBuddyControl_Timer()
If PropBuddyControlInit = False Then
    Me.BuddyControl = PropBuddyName
    PropBuddyControlInit = True
End If
TimerBuddyControl.Enabled = False
End Sub

Public Property Get ControlsEnum() As VBRUN.ParentControls
Attribute ControlsEnum.VB_MemberFlags = "40"
Set ControlsEnum = UserControl.ParentControls
End Property

Public Property Get ControlsExtender() As Object
Attribute ControlsExtender.VB_MemberFlags = "40"
Set ControlsExtender = Extender
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

Public Property Get Align() As Integer
Attribute Align.VB_Description = "Returns/sets a value that determines where an object is displayed on a form."
Attribute Align.VB_MemberFlags = "400"
Align = Extender.Align
End Property

Public Property Let Align(ByVal Value As Integer)
Extender.Align = Value
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
hWnd = PagerHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
PropBackColor = Value
If PagerHandle <> 0 Then SendMessage PagerHandle, PGM_SETBKCOLOR, 0, ByVal WinColor(PropBackColor)
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If PagerHandle <> 0 Then EnableWindow PagerHandle, IIf(Value = True, 1, 0)
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

Public Property Get OLEDragDropScroll() As Boolean
Attribute OLEDragDropScroll.VB_Description = "Returns/Sets whether this object will scroll during an OLE drag/drop operation."
OLEDragDropScroll = PropOLEDragDropScroll
End Property

Public Property Let OLEDragDropScroll(ByVal Value As Boolean)
PropOLEDragDropScroll = Value
If PagerHandle <> 0 Then Call ReCreatePager
UserControl.PropertyChanged "OLEDragDropScroll"
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
If PagerDesignMode = False Then Call RefreshMousePointer
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
        If PagerDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If PagerDesignMode = False Then Call RefreshMousePointer
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
If PagerDesignMode = False Then Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
If PagerHandle <> 0 Then Call ComCtlsSetRightToLeft(PagerHandle, dwMask)
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

Public Property Get BuddyControl() As Variant
Attribute BuddyControl.VB_Description = "Returns/sets the buddy control."
If PagerDesignMode = False Then
    If PropBuddyControlInit = False And PagerBuddyObjectPointer = 0 Then
        If Not PropBuddyName = "(None)" Then Me.BuddyControl = PropBuddyName
        PropBuddyControlInit = True
    End If
    Set BuddyControl = PropBuddyControl
Else
    BuddyControl = PropBuddyName
End If
End Property

Public Property Set BuddyControl(ByVal Value As Variant)
Me.BuddyControl = Value
End Property

Public Property Let BuddyControl(ByVal Value As Variant)
If PagerDesignMode = False Then
    If PagerHandle <> 0 Then
        Dim Success As Boolean, Handle As Long
        On Error Resume Next
        If IsObject(Value) Then
            If ControlIsValid(Value, Handle) = True Then
                Success = CBool(Handle <> 0)
                If Success = True Then
                    If PagerBuddyControlHandle <> 0 Then
                        SendMessage PagerHandle, PGM_SETCHILD, 0, ByVal 0&
                        SetParent PagerBuddyControlHandle, PagerBuddyControlPrevParent
                    End If
                    PagerBuddyControlHandle = Handle
                    PagerBuddyControlPrevParent = GetAncestor(Handle, GA_PARENT)
                    SetParent Handle, PagerHandle
                    SendMessage PagerHandle, PGM_SETCHILD, 0, ByVal Handle
                    PagerBuddyObjectPointer = ObjPtr(Value)
                    PropBuddyName = ProperControlName(Value)
                End If
            End If
        ElseIf VarType(Value) = vbString Then
            Dim ControlEnum As Object, CompareName As String
            For Each ControlEnum In UserControl.ParentControls
                If ControlIsValid(ControlEnum, Handle) = True Then
                    CompareName = ProperControlName(ControlEnum)
                    If CompareName = Value And Not CompareName = vbNullString Then
                        Success = CBool(Handle <> 0)
                        If Success = True Then
                            If PagerBuddyControlHandle <> 0 Then
                                SendMessage PagerHandle, PGM_SETCHILD, 0, ByVal 0&
                                SetParent PagerBuddyControlHandle, PagerBuddyControlPrevParent
                            End If
                            PagerBuddyControlHandle = Handle
                            PagerBuddyControlPrevParent = GetAncestor(Handle, GA_PARENT)
                            SetParent Handle, PagerHandle
                            SendMessage PagerHandle, PGM_SETCHILD, 0, ByVal Handle
                            PagerBuddyObjectPointer = ObjPtr(ControlEnum)
                            PropBuddyName = Value
                            Exit For
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            If PagerBuddyControlHandle <> 0 Then
                SendMessage PagerHandle, PGM_SETCHILD, 0, ByVal 0&
                SetParent PagerBuddyControlHandle, PagerBuddyControlPrevParent
                PagerBuddyControlHandle = 0
                PagerBuddyControlPrevParent = 0
            End If
            PagerBuddyObjectPointer = 0
            PropBuddyName = "(None)"
        End If
    End If
Else
    PropBuddyName = Value
End If
UserControl.PropertyChanged "BuddyControl"
End Property

Public Property Get Orientation() As PgrOrientationConstants
Attribute Orientation.VB_Description = "Returns/sets the orientation."
Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As PgrOrientationConstants)
Select Case Value
    Case PgrOrientationVertical, PgrOrientationHorizontal
        PropOrientation = Value
    Case Else
        Err.Raise 380
End Select
If PagerHandle <> 0 Then Call ReCreatePager
UserControl.PropertyChanged "Orientation"
End Property

Public Property Get BorderWidth() As Single
Attribute BorderWidth.VB_Description = "Returns/sets the border width that is drawn around the container."
BorderWidth = UserControl.ScaleX(PropBorderWidth, vbPixels, vbContainerSize)
End Property

Public Property Let BorderWidth(ByVal Value As Single)
If Value < 0 Then
    If PagerDesignMode = True Then
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
    PropBorderWidth = IntValue
    If PagerHandle <> 0 Then
        SendMessage PagerHandle, PGM_SETBORDER, 0, ByVal PropBorderWidth
        Me.Refresh
    End If
Else
    If PagerDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
UserControl.PropertyChanged "BorderWidth"
End Property

Public Property Get AutoScroll() As Boolean
Attribute AutoScroll.VB_Description = "Returns/sets a value that determines whether or not the pager control will automatically scroll the buddy control when the mouse hovers over a scroll button."
AutoScroll = PropAutoScroll
End Property

Public Property Let AutoScroll(ByVal Value As Boolean)
PropAutoScroll = Value
If PagerHandle <> 0 Then
    Dim dwStyle As Long
    dwStyle = GetWindowLong(PagerHandle, GWL_STYLE)
    If PropAutoScroll = True Then
        If Not (dwStyle And PGS_AUTOSCROLL) = PGS_AUTOSCROLL Then dwStyle = dwStyle Or PGS_AUTOSCROLL
    Else
        If (dwStyle And PGS_AUTOSCROLL) = PGS_AUTOSCROLL Then dwStyle = dwStyle And Not PGS_AUTOSCROLL
    End If
    SetWindowLong PagerHandle, GWL_STYLE, dwStyle
End If
UserControl.PropertyChanged "AutoScroll"
End Property

Public Property Get ButtonSize() As Single
Attribute ButtonSize.VB_Description = "Returns/sets the current button size. A value of -1 indicates that the default system size will be used."
If PagerHandle <> 0 And PagerDesignMode = False Then
    Select Case PropOrientation
        Case PgrOrientationVertical
            ButtonSize = UserControl.ScaleY(SendMessage(PagerHandle, PGM_GETBUTTONSIZE, 0, ByVal 0&), vbPixels, vbContainerSize)
        Case PgrOrientationHorizontal
            ButtonSize = UserControl.ScaleX(SendMessage(PagerHandle, PGM_GETBUTTONSIZE, 0, ByVal 0&), vbPixels, vbContainerSize)
    End Select
Else
    If PropButtonSize = -1 Then
        ButtonSize = -1
    Else
        Select Case PropOrientation
            Case PgrOrientationVertical
                ButtonSize = UserControl.ScaleY(PropButtonSize, vbPixels, vbContainerSize)
            Case PgrOrientationHorizontal
                ButtonSize = UserControl.ScaleX(PropButtonSize, vbPixels, vbContainerSize)
        End Select
    End If
End If
End Property

Public Property Let ButtonSize(ByVal Value As Single)
If Value < 0 And Not Value = -1 Then
    If PagerDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
Dim LngValue As Long, ErrValue As Long
On Error Resume Next
If Value = -1 Then
    LngValue = -1
Else
    Select Case PropOrientation
        Case PgrOrientationVertical
            LngValue = CLng(UserControl.ScaleY(Value, vbContainerSize, vbPixels))
        Case PgrOrientationHorizontal
            LngValue = CLng(UserControl.ScaleX(Value, vbContainerSize, vbPixels))
    End Select
End If
ErrValue = Err.Number
On Error GoTo 0
If (LngValue < 0 And Not LngValue = -1) Or ErrValue <> 0 Then
    If PagerDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
Else
    PropButtonSize = LngValue
    If PagerHandle <> 0 Then
        If PropButtonSize = -1 Then
            ' By default, the pager control sets its button size to three-fourths of the width of the scroll bar.
            Const SM_CXVSCROLL As Long = 2, SM_CYHSCROLL As Long = 3
            Select Case PropOrientation
                Case PgrOrientationVertical
                    LngValue = Fix(((GetSystemMetrics(SM_CXVSCROLL) / 4) * 3))
                Case PgrOrientationHorizontal
                    LngValue = Fix(((GetSystemMetrics(SM_CYHSCROLL) / 4) * 3))
            End Select
            SendMessage PagerHandle, PGM_SETBUTTONSIZE, 0, ByVal LngValue
        Else
            ' There is a minimum size to the pager button, currently 12 pixels.
            ' However, this can change so you should not depend on this value.
            SendMessage PagerHandle, PGM_SETBUTTONSIZE, 0, ByVal PropButtonSize
            PropButtonSize = SendMessage(PagerHandle, PGM_GETBUTTONSIZE, 0, ByVal 0&)
        End If
    End If
End If
UserControl.PropertyChanged "ButtonSize"
End Property

Private Sub CreatePager()
If PagerHandle <> 0 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If PropOLEDragDropScroll = True Then dwStyle = dwStyle Or PGS_DRAGNDROP
Select Case PropOrientation
    Case PgrOrientationVertical
        dwStyle = dwStyle Or PGS_VERT
    Case PgrOrientationHorizontal
        dwStyle = dwStyle Or PGS_HORZ
End Select
If PropAutoScroll = True Then dwStyle = dwStyle Or PGS_AUTOSCROLL
PagerHandle = CreateWindowEx(dwExStyle, StrPtr("SysPager"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If PagerHandle <> 0 Then
    SendMessage PagerHandle, PGM_FORWARDMOUSE, 1, ByVal 0&
    SendMessage PagerHandle, PGM_SETBORDER, 0, ByVal PropBorderWidth
    If Not PropButtonSize = -1 Then SendMessage PagerHandle, PGM_SETBUTTONSIZE, 0, ByVal PropButtonSize
End If
Me.BackColor = PropBackColor
If PagerDesignMode = False Then
    If PagerHandle <> 0 Then Call ComCtlsSetSubclass(PagerHandle, Me, 1)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
End Sub

Private Sub ReCreatePager()
If PagerDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    If PagerHandle <> 0 And PagerBuddyControlHandle <> 0 Then
        SendMessage PagerHandle, PGM_SETCHILD, 0, ByVal 0&
        SetParent PagerBuddyControlHandle, PagerBuddyControlPrevParent
        PagerBuddyControlHandle = 0
        PagerBuddyControlPrevParent = 0
    End If
    Call DestroyPager
    Call CreatePager
    Call UserControl_Resize
    If Not PropBuddyControl Is Nothing Then Set Me.BuddyControl = PropBuddyControl
    If Locked = True Then LockWindowUpdate 0
    Me.Refresh
Else
    Call DestroyPager
    Call CreatePager
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyPager()
If PagerHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(PagerHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
If PagerBuddyControlHandle <> 0 Then
    SendMessage PagerHandle, PGM_SETCHILD, 0, ByVal 0&
    SetParent PagerBuddyControlHandle, PagerBuddyControlPrevParent
    PagerBuddyControlHandle = 0
    PagerBuddyControlPrevParent = 0
End If
ShowWindow PagerHandle, SW_HIDE
SetParent PagerHandle, 0
DestroyWindow PagerHandle
PagerHandle = 0
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub ReCalcSize()
Attribute ReCalcSize.VB_Description = "Forces the pager control to recalculate the size of the buddy control."
If PagerHandle <> 0 Then SendMessage PagerHandle, PGM_RECALCSIZE, 0, ByVal 0&
End Sub

Public Function GetButtonState(ByVal Button As PgrButtonConstants) As PgrButtonStateConstants
Attribute GetButtonState.VB_Description = "Retrieves the state of the specified button. The return value can be a combination of values using a bitwise OR."
Select Case Button
    Case PgrButtonTopOrLeft, PgrButtonBottomOrRight
        If PagerHandle <> 0 Then GetButtonState = SendMessage(PagerHandle, PGM_GETBUTTONSTATE, 0, ByVal CLng(Button))
    Case Else
        Err.Raise 380
End Select
End Function

Public Property Get Value() As Single
Attribute Value.VB_Description = "Returns/sets the current position."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "400"
If PagerHandle <> 0 Then
    Select Case PropOrientation
        Case PgrOrientationVertical
            Value = UserControl.ScaleY(SendMessage(PagerHandle, PGM_GETPOS, 0, ByVal 0&), vbPixels, vbContainerPosition)
        Case PgrOrientationHorizontal
            Value = UserControl.ScaleX(SendMessage(PagerHandle, PGM_GETPOS, 0, ByVal 0&), vbPixels, vbContainerPosition)
    End Select
End If
End Property

Public Property Let Value(ByVal NewValue As Single)
If NewValue < 0 Then Err.Raise 380
Dim LngValue As Long
On Error Resume Next
Select Case PropOrientation
    Case PgrOrientationVertical
        LngValue = CLng(UserControl.ScaleY(NewValue, vbContainerPosition, vbPixels))
    Case PgrOrientationHorizontal
        LngValue = CLng(UserControl.ScaleX(NewValue, vbContainerPosition, vbPixels))
End Select
If Err.Number <> 0 Then LngValue = 0
On Error GoTo 0
If LngValue >= 0 Then
    If PagerHandle <> 0 Then SendMessage PagerHandle, PGM_SETPOS, 0, ByVal LngValue
Else
    Err.Raise 380
End If
End Property

Private Function ControlIsValid(ByVal Control As Object, ByRef Handle As Long) As Boolean
On Error Resume Next
Dim Container As Object
Set Container = Control.Container
ControlIsValid = CBool(Err.Number = 0 And Not Control Is Extender)
On Error GoTo 0
Handle = 0
If ControlIsValid = True Then
    On Error Resume Next
    Handle = Control.hWndUserControl
    If Err.Number <> 0 Then Handle = Control.hWnd
    On Error GoTo 0
End If
End Function

Private Function PropBuddyControl() As Object
If PagerBuddyObjectPointer <> 0 Then Set PropBuddyControl = PtrToObj(PagerBuddyObjectPointer)
End Function

Private Function ISubclass_Message(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
Select Case dwRefData
    Case 1
        ISubclass_Message = WindowProcControl(hWnd, wMsg, wParam, lParam)
    Case 2
        ISubclass_Message = WindowProcUserControl(hWnd, wMsg, wParam, lParam)
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
                PagerIsClick = True
            Case WM_MBUTTONDOWN
                RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
                PagerIsClick = True
            Case WM_RBUTTONDOWN
                RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
                PagerIsClick = True
            Case WM_MOUSEMOVE
                If PagerMouseOver = False And PropMouseTrack = True Then
                    PagerMouseOver = True
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
                If PagerIsClick = True Then
                    PagerIsClick = False
                    If (X >= 0 And X <= UserControl.Width) And (Y >= 0 And Y <= UserControl.Height) Then RaiseEvent Click
                End If
        End Select
    Case WM_MOUSELEAVE
        If PagerMouseOver = True Then
            PagerMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = PagerHandle Then
            Select Case NM.Code
                Case PGN_SCROLL
                    Dim NMPGS As NMPGSCROLL
                    CopyMemory NMPGS, ByVal lParam, LenB(NMPGS)
                    Dim Delta As Single, RC As RECT
                    CopyMemory RC, NMPGS.RCParent(0), 16
                    With UserControl
                    Select Case NMPGS.iDir(0)
                        Case PGF_SCROLLUP, PGF_SCROLLDOWN
                            Delta = .ScaleY(NMPGS.iScroll(0), vbPixels, vbContainerPosition)
                            RaiseEvent Scroll(NMPGS.fwKeys, NMPGS.iDir(0), 0, .ScaleY(NMPGS.iYPos(0), vbPixels, vbContainerPosition), Delta, .ScaleX(RC.Left, vbPixels, vbContainerPosition), .ScaleY(RC.Top, vbPixels, vbContainerPosition), .ScaleX(RC.Right, vbPixels, vbContainerPosition), .ScaleY(RC.Bottom, vbPixels, vbContainerPosition))
                            NMPGS.iScroll(0) = CInt(UserControl.ScaleY(Delta, vbContainerPosition, vbPixels))
                        Case PGF_SCROLLLEFT, PGF_SCROLLRIGHT
                            Delta = .ScaleX(NMPGS.iScroll(0), vbPixels, vbContainerPosition)
                            RaiseEvent Scroll(NMPGS.fwKeys, NMPGS.iDir(0), .ScaleX(NMPGS.iXPos(0), vbPixels, vbContainerPosition), 0, Delta, .ScaleX(RC.Left, vbPixels, vbContainerPosition), .ScaleY(RC.Top, vbPixels, vbContainerPosition), .ScaleX(RC.Right, vbPixels, vbContainerPosition), .ScaleY(RC.Bottom, vbPixels, vbContainerPosition))
                            NMPGS.iScroll(0) = CInt(UserControl.ScaleY(Delta, vbContainerPosition, vbPixels))
                    End Select
                    End With
                    CopyMemory ByVal lParam, NMPGS, LenB(NMPGS)
                Case PGN_CALCSIZE
                    Dim NMPGCS As NMPGCALCSIZE
                    CopyMemory NMPGCS, ByVal lParam, LenB(NMPGCS)
                    Dim Size As Single
                    With NMPGCS
                    Select Case .dwFlag
                        Case PGF_CALCWIDTH
                            If Not Me.BuddyControl Is Nothing Then
                                Size = PropBuddyControl.Width
                            Else
                                Size = UserControl.ScaleX(.iWidth, vbPixels, vbContainerSize)
                            End If
                            RaiseEvent CalcSize(Size, 0)
                            .iWidth = CLng(UserControl.ScaleX(Size, vbContainerSize, vbPixels))
                        Case PGF_CALCHEIGHT
                            If Not Me.BuddyControl Is Nothing Then
                                Size = PropBuddyControl.Height
                            Else
                                Size = UserControl.ScaleY(.iHeight, vbPixels, vbContainerSize)
                            End If
                            RaiseEvent CalcSize(0, Size)
                            .iHeight = CLng(UserControl.ScaleY(Size, vbContainerSize, vbPixels))
                    End Select
                    End With
                    CopyMemory ByVal lParam, NMPGCS, LenB(NMPGCS)
                Case PGN_HOTITEMCHANGE
                    Dim NMPGHI As NMPGHOTITEM
                    CopyMemory NMPGHI, ByVal lParam, LenB(NMPGHI)
                    With NMPGHI
                    ' HICF_MOUSE is not included in dwFlags.
                    If (.dwFlags And HICF_ENTERING) = HICF_ENTERING Or (.dwFlags And HICF_LEAVING) = HICF_LEAVING Then
                        ' This check is necessary as else this would be fired repeatedly
                        ' when HICF_ENTERING is included in dwFlags.
                        If .dwFlags <> PagerHotItemChangePrevFlags Then
                            RaiseEvent HotChanged
                            PagerHotItemChangePrevFlags = .dwFlags
                        End If
                    End If
                    End With
            End Select
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function
