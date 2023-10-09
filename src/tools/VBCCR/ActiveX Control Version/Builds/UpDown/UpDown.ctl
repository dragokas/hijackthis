VERSION 5.00
Begin VB.UserControl UpDown 
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "UpDown.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "UpDown.ctx":0028
   Begin VB.Timer TimerBuddyControl 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "UpDown"
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
Private UdnOrientationVertical, UdnOrientationHorizontal
Private UdnNumberStyleDecimal, UdnNumberStyleHexadecimal
#End If
Public Enum UdnOrientationConstants
UdnOrientationVertical = 0
UdnOrientationHorizontal = 1
End Enum
Public Enum UdnNumberStyleConstants
UdnNumberStyleDecimal = 0
UdnNumberStyleHexadecimal = 1
End Enum
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type UDACCEL
nSec As Long
nInc As Long
End Type
Private Type NMHDR
hWndFrom As LongPtr
IDFrom As LongPtr
Code As Long
End Type
Private Type NMUPDOWN
hdr As NMHDR
iPos As Long
iDelta As Long
End Type
Public Event DownClick()
Attribute DownClick.VB_Description = "Occurs when the position has changed by a down click."
Public Event UpClick()
Attribute UpClick.VB_Description = "Occurs when the position has changed by an up click."
Public Event BeforeChange(ByVal Value As Long, ByRef Delta As Long)
Attribute BeforeChange.VB_Description = "Occurs when the position is about to change."
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
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByRef lpParam As Any) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
#End If
Private Const ICC_UPDOWN_CLASS As Long = &H10
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const SW_HIDE As Long = &H0
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115
Private Const WM_NOTIFY As Long = &H4E
Private Const UDN_FIRST As Long = (-721)
Private Const UDN_DELTAPOS As Long = (UDN_FIRST - 1)
Private Const UDS_WRAP As Long = &H1
Private Const UDS_HORZ As Long = &H40
Private Const UDS_HOTTRACK As Long = &H100
Private Const WM_USER As Long = &H400
Private Const UDM_SETRANGE As Long = (WM_USER + 101) ' 16 bit
Private Const UDM_GETRANGE As Long = (WM_USER + 102) ' 16 bit
Private Const UDM_SETRANGE32 As Long = (WM_USER + 111)
Private Const UDM_GETRANGE32 As Long = (WM_USER + 112)
Private Const UDM_SETPOS As Long = (WM_USER + 103) ' 16 bit
Private Const UDM_GETPOS As Long = (WM_USER + 104) ' 16 bit
Private Const UDM_SETPOS32 As Long = (WM_USER + 113)
Private Const UDM_GETPOS32 As Long = (WM_USER + 114)
Private Const UDM_SETACCEL As Long = (WM_USER + 107)
Private Const UDM_GETACCEL As Long = (WM_USER + 108)
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETUNICODEFORMAT As Long = (CCM_FIRST + 5)
Private Const UDM_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IPerPropertyBrowsingVB
Private UpDownHandle As LongPtr
Private UpDownMouseOver As Boolean
Private UpDownDesignMode As Boolean
Private UpDownBuddyObjectPointer As LongPtr
Private DispIDBuddyControl As Long, BuddyControlArray() As String
Private PropVisualStyles As Boolean
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropBuddyName As String, PropBuddyControlInit As Boolean
Private PropBuddyProperty As String
Private PropMin As Long, PropMax As Long
Private PropValue As Long, PropIncrement As Long
Private PropWrap As Boolean
Private PropHotTracking As Boolean
Private PropOrientation As UdnOrientationConstants
Private PropThousandsSeparator As Boolean
Private PropNumberStyle As UdnNumberStyleConstants

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByRef pdwSupportedOptions As Long, ByRef pdwEnabledOptions As Long)
Const INTERFACESAFE_FOR_UNTRUSTED_CALLER As Long = &H1, INTERFACESAFE_FOR_UNTRUSTED_DATA As Long = &H2
pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByRef riid As OLEGuids.OLECLSID, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
End Sub

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
If DispID = DispIDBuddyControl Then
    DisplayName = PropBuddyName
    Handled = True
End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
If DispID = DispIDBuddyControl Then
    On Error GoTo CATCH_EXCEPTION
    Dim ControlEnum As Object, PropUBound As Long
    PropUBound = UBound(StringsOut())
    ReDim Preserve StringsOut(PropUBound + 1) As String
    ReDim Preserve CookiesOut(PropUBound + 1) As Long
    StringsOut(PropUBound) = "(None)"
    CookiesOut(PropUBound) = PropUBound
    For Each ControlEnum In UserControl.ParentControls
        If ControlIsValid(ControlEnum) = True Then
            If ControlEnum.Container Is Extender.Container Then
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
If DispID = DispIDBuddyControl Then
    If Cookie < UBound(BuddyControlArray()) Then Value = BuddyControlArray(Cookie)
    Handled = True
End If
End Sub

Private Sub UserControl_Initialize()
Call ComCtlsLoadShellMod
Call ComCtlsInitCC(ICC_UPDOWN_CLASS)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim BuddyControlArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIDBuddyControl = 0 Then DispIDBuddyControl = GetDispID(Me, "BuddyControl")
On Error Resume Next
UpDownDesignMode = Not Ambient.UserMode
On Error GoTo 0
PropVisualStyles = True
Me.OLEDropMode = vbOLEDropNone
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropBuddyName = "(None)"
PropBuddyProperty = vbNullString
PropMin = 0
PropMax = 10
PropValue = 0
PropIncrement = 1
PropWrap = False
PropHotTracking = True
PropOrientation = UdnOrientationVertical
PropThousandsSeparator = True
PropNumberStyle = UdnNumberStyleDecimal
Call CreateUpDown
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDBuddyControl = 0 Then DispIDBuddyControl = GetDispID(Me, "BuddyControl")
On Error Resume Next
UpDownDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropBuddyName = .ReadProperty("BuddyControl", "(None)")
PropBuddyProperty = VarToStr(.ReadProperty("BuddyProperty", vbNullString))
PropMin = .ReadProperty("Min", 0)
PropMax = .ReadProperty("Max", 10)
PropValue = .ReadProperty("Value", 0)
PropIncrement = .ReadProperty("Increment", 1)
PropWrap = .ReadProperty("Wrap", False)
PropHotTracking = .ReadProperty("HotTracking", True)
PropOrientation = .ReadProperty("Orientation", UdnOrientationVertical)
PropThousandsSeparator = .ReadProperty("ThousandsSeparator", True)
PropNumberStyle = .ReadProperty("NumberStyle", UdnNumberStyleDecimal)
End With
Call CreateUpDown
If Not PropBuddyName = "(None)" Then TimerBuddyControl.Enabled = Ambient.UserMode
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "BuddyControl", PropBuddyName, "(None)"
.WriteProperty "BuddyProperty", StrToVar(PropBuddyProperty), vbNullString
.WriteProperty "Min", PropMin, 0
.WriteProperty "Max", PropMax, 10
.WriteProperty "Value", PropValue, 0
.WriteProperty "Increment", PropIncrement, 1
.WriteProperty "Wrap", PropWrap, False
.WriteProperty "HotTracking", PropHotTracking, True
.WriteProperty "Orientation", PropOrientation, UdnOrientationVertical
.WriteProperty "ThousandsSeparator", PropThousandsSeparator, True
.WriteProperty "NumberStyle", PropNumberStyle, UdnNumberStyleDecimal
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
If UpDownHandle = NULL_PTR Then InProc = False: Exit Sub
Dim WndRect As RECT
GetWindowRect UpDownHandle, WndRect
Select Case PropOrientation
    Case UdnOrientationHorizontal
        MoveWindow UpDownHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    Case UdnOrientationVertical
        MoveWindow UpDownHandle, 0, 0, (WndRect.Right - WndRect.Left), .ScaleHeight, 1
        .Extender.Width = .ScaleX((WndRect.Right - WndRect.Left), vbPixels, vbContainerSize)
End Select
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyUpDown
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

Public Property Get ControlsContainer() As Object
Attribute ControlsContainer.VB_MemberFlags = "40"
Set ControlsContainer = Extender.Container
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
hWnd = UpDownHandle
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

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not. Requires comctl32.dll version 6.0 or higher."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If UpDownHandle <> NULL_PTR And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles UpDownHandle
    Else
        RemoveVisualStyles UpDownHandle
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
If UpDownHandle <> NULL_PTR Then EnableWindow UpDownHandle, IIf(Value = True, 1, 0)
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
If UpDownDesignMode = False Then Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
If UpDownHandle <> NULL_PTR Then Call ComCtlsSetRightToLeft(UpDownHandle, dwMask)
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
If UpDownDesignMode = False Then
    If PropBuddyControlInit = False And PropBuddyControl Is Nothing Then
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
If UpDownDesignMode = False Then
    If UpDownHandle <> NULL_PTR Then
        Dim Success As Boolean
        On Error Resume Next
        If IsObject(Value) Then
            If ControlIsValid(Value) = True Then
                If Value.Container Is Extender.Container Then
                    Success = CBool(Err.Number = 0)
                    If Success = True Then
                        UpDownBuddyObjectPointer = ObjPtr(Value)
                        If ProperControlName(Value) <> PropBuddyName Then PropBuddyProperty = vbNullString
                        PropBuddyName = ProperControlName(Value)
                    End If
                End If
            End If
        ElseIf VarType(Value) = vbString Then
            Dim ControlEnum As Object, CompareName As String
            For Each ControlEnum In UserControl.ParentControls
                If ControlIsValid(ControlEnum) = True Then
                    If ControlEnum.Container Is Extender.Container Then
                        CompareName = ProperControlName(ControlEnum)
                        If CompareName = Value And Not CompareName = vbNullString Then
                            Err.Clear
                            Success = CBool(Err.Number = 0)
                            If Success = True Then
                                UpDownBuddyObjectPointer = ObjPtr(ControlEnum)
                                If Value <> PropBuddyName Then PropBuddyProperty = vbNullString
                                PropBuddyName = Value
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next ControlEnum
        End If
        On Error GoTo 0
        If Success = False Then
            UpDownBuddyObjectPointer = NULL_PTR
            PropBuddyName = "(None)"
            PropBuddyProperty = vbNullString
        End If
    End If
Else
    If Value <> PropBuddyName Or Value = "(None)" Then PropBuddyProperty = vbNullString
    PropBuddyName = Value
End If
UserControl.PropertyChanged "BuddyControl"
End Property

Public Property Get BuddyProperty() As String
Attribute BuddyProperty.VB_Description = "Returns/sets the property used to synchronize the control with its buddy control. Specifying a property named '_Value' means that the default property will be used."
Attribute BuddyProperty.VB_ProcData.VB_Invoke_Property = "PPUpDownGeneral"
BuddyProperty = PropBuddyProperty
End Property

Public Property Let BuddyProperty(ByVal Value As String)
Dim BuddyControl As Control
Set BuddyControl = GetBuddyControl()
If Not BuddyControl Is Nothing Then
    Dim Success As Boolean
    On Error Resume Next
    Select Case Value
        Case vbNullString, "_Value"
            Success = True
        Case Else
            CallByName BuddyControl, Value, VbLet, CallByName(BuddyControl, Value, VbGet)
            Success = CBool(Err.Number = 0)
    End Select
    On Error GoTo 0
    If Success = False Then
        If UpDownDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            PropBuddyProperty = vbNullString
            Exit Property
        Else
            Err.Raise 380
        End If
    Else
        PropBuddyProperty = Value
    End If
Else
    Select Case Value
        Case vbNullString
            PropBuddyProperty = Value
        Case Else
            If UpDownDesignMode = True Then
                MsgBox "BuddyControl property must be set first", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise Number:=35754, Description:="BuddyControl property must be set first"
            End If
    End Select
End If
UserControl.PropertyChanged "BuddyProperty"
End Property

Public Property Get SyncBuddy() As Boolean
Attribute SyncBuddy.VB_Description = "Returns/sets a value indicating if the control synchronizes the value property with a property in the buddy control. In order to synchronize changes done from the buddy control it is necessary to use the 'SyncFromBuddy' method."
SyncBuddy = Not CBool(PropBuddyProperty = vbNullString)
End Property

Public Property Let SyncBuddy(ByVal Value As Boolean)
If Value = True Then
    If GetBuddyControl() Is Nothing Then
        If UpDownDesignMode = True Then
            MsgBox "BuddyControl property must be set first", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise Number:=35754, Description:="BuddyControl property must be set first"
        End If
    End If
    If PropBuddyProperty = vbNullString Then PropBuddyProperty = "_Value"
Else
    PropBuddyProperty = vbNullString
End If
UserControl.PropertyChanged "SyncBuddy"
End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum value."
If UpDownHandle <> NULL_PTR Then
    SendMessage UpDownHandle, UDM_GETRANGE32, VarPtr(Min), ByVal 0&
Else
    Min = PropMin
End If
End Property

Public Property Let Min(ByVal Value As Long)
If Value <= Me.Max Then
    PropMin = Value
    If Me.Value < PropMin Then Me.Value = PropMin
Else
    If UpDownDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If UpDownHandle <> NULL_PTR Then SendMessage UpDownHandle, UDM_SETRANGE32, PropMin, ByVal PropMax
Me.Refresh
UserControl.PropertyChanged "Min"
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum value."
If UpDownHandle <> NULL_PTR Then
    SendMessage UpDownHandle, UDM_GETRANGE32, 0, ByVal VarPtr(Max)
Else
    Max = PropMax
End If
End Property

Public Property Let Max(ByVal Value As Long)
If Value >= Me.Min Then
    PropMax = Value
    If Me.Value > PropMax Then Me.Value = PropMax
Else
    If UpDownDesignMode = True Then
        MsgBox "Invalid property value", vbCritical + vbOKOnly
        Exit Property
    Else
        Err.Raise 380
    End If
End If
If UpDownHandle <> NULL_PTR Then SendMessage UpDownHandle, UDM_SETRANGE32, PropMin, ByVal PropMax
Me.Refresh
UserControl.PropertyChanged "Max"
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the current position."
Attribute Value.VB_UserMemId = 0
If UpDownHandle <> NULL_PTR Then
    Value = CLng(SendMessage(UpDownHandle, UDM_GETPOS32, 0, ByVal 0&))
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
If UpDownHandle <> NULL_PTR Then SendMessage UpDownHandle, UDM_SETPOS32, 0, ByVal PropValue
UserControl.PropertyChanged "Value"
Call SyncProperty(False)
If Changed = True Then RaiseEvent Change
End Property

Public Property Get Increment() As Long
Attribute Increment.VB_Description = "Returns/sets the position change increment."
If UpDownHandle <> NULL_PTR Then
    Dim Accel As UDACCEL
    SendMessage UpDownHandle, UDM_GETACCEL, 1, Accel
    Increment = Accel.nInc
Else
    Increment = PropIncrement
End If
End Property

Public Property Let Increment(ByVal Value As Long)
PropIncrement = Value
If UpDownHandle <> NULL_PTR Then
    Dim Accel As UDACCEL
    Accel.nSec = 0
    Accel.nInc = PropIncrement
    SendMessage UpDownHandle, UDM_SETACCEL, 1, Accel
End If
UserControl.PropertyChanged "Increment"
End Property

Public Property Get Wrap() As Boolean
Attribute Wrap.VB_Description = "Returns/sets a value that determines whether or not the position will be wrapped if it is incremented or decremented beyond the ending or beginning of the range."
Wrap = PropWrap
End Property

Public Property Let Wrap(ByVal Value As Boolean)
PropWrap = Value
If UpDownHandle <> NULL_PTR Then Call ReCreateUpDown
UserControl.PropertyChanged "Wrap"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets a value that determines whether or not the control highlights the up arrow and down arrow as the pointer passes over them. This flag is ignored on Windows XP (or above) when the desktop theme overrides it."
HotTracking = PropHotTracking
End Property

Public Property Let HotTracking(ByVal Value As Boolean)
PropHotTracking = Value
If UpDownHandle <> NULL_PTR Then Call ReCreateUpDown
UserControl.PropertyChanged "HotTracking"
End Property

Public Property Get Orientation() As UdnOrientationConstants
Attribute Orientation.VB_Description = "Returns/sets the orientation."
Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As UdnOrientationConstants)
Select Case Value
    Case UdnOrientationVertical, UdnOrientationHorizontal
        Dim SwapSizes As Boolean
        SwapSizes = CBool(PropOrientation <> Value)
        PropOrientation = Value
        If SwapSizes = True Then
            With UserControl
            .Extender.Move .Extender.Left, .Extender.Top, .Extender.Height, .Extender.Width
            End With
        End If
    Case Else
        Err.Raise 380
End Select
If UpDownHandle <> NULL_PTR Then Call ReCreateUpDown
UserControl.PropertyChanged "Orientation"
End Property

Public Property Get ThousandsSeparator() As Boolean
Attribute ThousandsSeparator.VB_Description = "Returns/sets a value indicating if a thousands separator is inserted between every three decimal digits. Only applicable if the number style property is set to decimal."
ThousandsSeparator = PropThousandsSeparator
End Property

Public Property Let ThousandsSeparator(ByVal Value As Boolean)
PropThousandsSeparator = Value
UserControl.PropertyChanged "ThousandsSeparator"
End Property

Public Property Get NumberStyle() As UdnNumberStyleConstants
Attribute NumberStyle.VB_Description = "Returns/sets the number style."
NumberStyle = PropNumberStyle
End Property

Public Property Let NumberStyle(ByVal Value As UdnNumberStyleConstants)
Select Case Value
    Case UdnNumberStyleDecimal, UdnNumberStyleHexadecimal
        PropNumberStyle = Value
    Case Else
        Err.Raise 380
End Select
UserControl.PropertyChanged "NumberStyle"
End Property

Private Sub CreateUpDown()
If UpDownHandle <> NULL_PTR Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE
If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
If PropWrap = True Then dwStyle = dwStyle Or UDS_WRAP
If PropHotTracking = True Then dwStyle = dwStyle Or UDS_HOTTRACK
If PropOrientation = UdnOrientationHorizontal Then dwStyle = dwStyle Or UDS_HORZ
UpDownHandle = CreateWindowEx(dwExStyle, StrPtr("msctls_updown32"), NULL_PTR, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, NULL_PTR, App.hInstance, ByVal NULL_PTR)
If UpDownHandle <> NULL_PTR Then
    SendMessage UpDownHandle, UDM_SETUNICODEFORMAT, 1, ByVal 0&
    SendMessage UpDownHandle, UDM_SETRANGE32, PropMin, ByVal PropMax
End If
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Value = PropValue
Me.Increment = PropIncrement
If UpDownDesignMode = False Then
    If UpDownHandle <> NULL_PTR Then Call ComCtlsSetSubclass(UpDownHandle, Me, 1)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
End Sub

Private Sub ReCreateUpDown()
If UpDownDesignMode = False Then
    Dim Locked As Boolean
    Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
    Call DestroyUpDown
    Call CreateUpDown
    Call UserControl_Resize
    If Not PropBuddyControl Is Nothing Then Set Me.BuddyControl = PropBuddyControl
    If Locked = True Then LockWindowUpdate NULL_PTR
    Me.Refresh
Else
    Call DestroyUpDown
    Call CreateUpDown
    Call UserControl_Resize
End If
End Sub

Private Sub DestroyUpDown()
If UpDownHandle = NULL_PTR Then Exit Sub
Call ComCtlsRemoveSubclass(UpDownHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow UpDownHandle, SW_HIDE
SetParent UpDownHandle, NULL_PTR
DestroyWindow UpDownHandle
UpDownHandle = NULL_PTR
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
UserControl.Refresh
RedrawWindow UserControl.hWnd, NULL_PTR, NULL_PTR, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub SyncFromBuddy()
Attribute SyncFromBuddy.VB_Description = "Method to update the value property with the property in the buddy control."
Call SyncProperty(True)
End Sub

Private Sub SyncProperty(Optional ByVal FromBuddy As Boolean)
Static InProc As Boolean
If UpDownHandle = NULL_PTR Or PropBuddyProperty = vbNullString Or InProc = True Then Exit Sub
If UpDownDesignMode = False Then
    Dim VarValue As Variant, LngValue As Long
    If Not PropBuddyControl Is Nothing Then
        On Error Resume Next
        If PropBuddyProperty = "_Value" Then
            VarValue = CallByName(PropBuddyControl, "", VbGet)
        Else
            VarValue = CallByName(PropBuddyControl, PropBuddyProperty, VbGet)
        End If
        If FromBuddy = True Then
            Select Case VarType(VarValue)
                Case vbString
                    If PropThousandsSeparator = True And PropNumberStyle = UdnNumberStyleDecimal Then
                        Dim GroupDigit As String
                        GroupDigit = Mid$(VBA.FormatNumber(1000, 0, , , vbTrue), 2, 1)
                        If Not (GroupDigit = vbNullString Or GroupDigit = "0") Then
                            LngValue = CLng(Replace(VarValue, GroupDigit, vbNullString))
                        Else
                            LngValue = CLng(VarValue)
                        End If
                    Else
                        If Left(VarValue, 2) = "0x" And PropNumberStyle = UdnNumberStyleHexadecimal Then
                            LngValue = CLng("&H" & Mid(VarValue, 3))
                        Else
                            LngValue = CLng(VarValue)
                        End If
                    End If
                Case vbLong, vbInteger, vbByte
                    LngValue = VarValue
                Case vbDouble, vbSingle
                    LngValue = CLng(VarValue)
            End Select
            If Err.Number = 0 Then
                If LngValue <> Me.Value Then
                    InProc = True
                    Me.Value = LngValue
                    InProc = False
                End If
            End If
        Else
            Select Case VarType(VarValue)
                Case vbString
                    Dim StrValue As String
                    If PropThousandsSeparator = True And PropNumberStyle = UdnNumberStyleDecimal Then
                        StrValue = Format$(Me.Value, "#,###,###,##0")
                    Else
                        If PropNumberStyle = UdnNumberStyleHexadecimal Then
                            LngValue = Me.Value
                            StrValue = "0x" & String$(IIf(Len(Hex$(LngValue)) < 5, 4, 8) - Len(Hex$(LngValue)), "0") & Hex$(LngValue)
                        Else
                            StrValue = CStr(Me.Value)
                        End If
                    End If
                    If PropBuddyProperty = "_Value" Then
                        CallByName PropBuddyControl, "", VbLet, StrValue
                    ElseIf Not PropBuddyProperty = vbNullString Then
                        CallByName PropBuddyControl, PropBuddyProperty, VbLet, StrValue
                    End If
                Case vbDouble, vbSingle, vbLong, vbInteger, vbByte
                    If PropBuddyProperty = "_Value" Then
                        CallByName PropBuddyControl, "", VbLet, Me.Value
                    ElseIf Not PropBuddyProperty = vbNullString Then
                        CallByName PropBuddyControl, PropBuddyProperty, VbLet, Me.Value
                    End If
            End Select
        End If
        On Error GoTo 0
    End If
End If
End Sub

Private Function GetBuddyControl() As Object
If UpDownDesignMode = False Then
    Set GetBuddyControl = PropBuddyControl
ElseIf Not PropBuddyName = "(None)" Then
    Dim ControlEnum As Object, CompareName As String, Success As Boolean
    For Each ControlEnum In UserControl.ParentControls
        If ControlIsValid(ControlEnum) = True Then
            If ControlEnum.Container Is Extender.Container Then
                CompareName = ProperControlName(ControlEnum)
                If CompareName = PropBuddyName And Not CompareName = vbNullString Then
                    Err.Clear
                    Success = CBool(Err.Number = 0)
                    If Success = True Then
                        Set GetBuddyControl = ControlEnum
                        Exit For
                    End If
                End If
            End If
        End If
    Next ControlEnum
End If
End Function

Private Function ControlIsValid(ByVal Control As Object) As Boolean
On Error Resume Next
Dim Container As Object
Set Container = Control.Container
ControlIsValid = CBool(Err.Number = 0 And Not Control Is Extender)
On Error GoTo 0
End Function

Private Function PropBuddyControl() As Object
If UpDownBuddyObjectPointer <> NULL_PTR Then Set PropBuddyControl = PtrToObj(UpDownBuddyObjectPointer)
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
End Select
End Function

Private Function WindowProcControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
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
                If UpDownMouseOver = False And PropMouseTrack = True Then
                    UpDownMouseOver = True
                    RaiseEvent MouseEnter
                    Call ComCtlsRequestMouseLeave(hWnd)
                End If
                RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), X, Y)
            Case WM_LBUTTONUP
                RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_MBUTTONUP
                RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), X, Y)
            Case WM_RBUTTONUP
                RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), X, Y)
        End Select
    Case WM_MOUSELEAVE
        If UpDownMouseOver = True Then
            UpDownMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Select Case wMsg
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = UpDownHandle Then
            If NM.Code = UDN_DELTAPOS Then
                Dim NMUD As NMUPDOWN
                CopyMemory NMUD, ByVal lParam, LenB(NMUD)
                RaiseEvent BeforeChange(NMUD.iPos, NMUD.iDelta)
                Select Case NMUD.iDelta
                    Case 0
                        WindowProcUserControl = 1
                        Exit Function
                    Case Is < 0
                        RaiseEvent DownClick
                    Case Is > 0
                        RaiseEvent UpClick
                End Select
            End If
        End If
    Case WM_VSCROLL, WM_HSCROLL
        If lParam = UpDownHandle Then
            Call SyncProperty(False)
            Dim NewValue As Long
            NewValue = CLng(SendMessage(UpDownHandle, UDM_GETPOS32, 0, ByVal 0&))
            If PropValue <> NewValue Then
                PropValue = NewValue
                RaiseEvent Change
            End If
        End If
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
End Function
