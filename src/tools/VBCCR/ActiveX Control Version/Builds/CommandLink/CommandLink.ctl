VERSION 5.00
Begin VB.UserControl CommandLink 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DefaultCancel   =   -1  'True
   DrawStyle       =   5  'Transparent
   HasDC           =   0   'False
   PropertyPages   =   "CommandLink.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "CommandLink.ctx":004C
   Begin VB.Timer TimerImageList 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "CommandLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type TACCEL
FVirt As Byte
Key As Integer
Cmd As Integer
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
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event HotChanged()
Attribute HotChanged.VB_Description = "Occurrs when the command link control's hot state changes."
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
Private Declare Function CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableW" (ByVal lpAccel As Long, ByVal cEntries As Long) As Long
Private Declare Function DestroyAcceleratorTable Lib "user32" (ByVal hAccel As Long) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanW" (ByVal cChar As Integer) As Integer
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const HWND_DESKTOP As Long = &H0
Private Const FVIRTKEY As Long = &H1
Private Const FSHIFT As Long = &H4
Private Const FALT As Long = &H10
Private Const GWL_STYLE As Long = (-16)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000
Private Const SW_HIDE As Long = &H0
Private Const GA_ROOT As Long = 2
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
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_COMMAND As Long = &H111
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_CTLCOLORBTN As Long = &H135
Private Const WM_PAINT As Long = &HF
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const BS_TEXT As Long = &H0
Private Const BS_PUSHBUTTON As Long = &H0
Private Const BS_DEFPUSHBUTTON As Long = &H1
Private Const BS_COMMANDLINK As Long = &HE
Private Const BS_ICON As Long = &H40
Private Const BS_BITMAP As Long = &H80
Private Const BS_NOTIFY As Long = &H4000
Private Const BM_GETSTATE As Long = &HF2
Private Const BM_SETSTATE As Long = &HF3
Private Const BM_SETIMAGE As Long = &HF7
Private Const BM_CLICK As Long = &HF5
Private Const BCM_FIRST As Long = &H1600
Private Const BCM_GETIDEALSIZE As Long = (BCM_FIRST + 1)
Private Const BCM_SETIMAGELIST As Long = (BCM_FIRST + 2)
Private Const BCM_GETIMAGELIST As Long = (BCM_FIRST + 3)
Private Const BCM_SETNOTE As Long = (BCM_FIRST + 9)
Private Const BCM_GETNOTE As Long = (BCM_FIRST + 10)
Private Const BCM_GETNOTELENGTH As Long = (BCM_FIRST + 11)
Private Const BCM_SETSHIELD As Long = (BCM_FIRST + 12)
Private Const BST_PUSHED As Long = &H4
Private Const BST_HOT As Long = &H200
Private Const BCCL_NOGLYPH As Long = (-1)
Private Const BN_CLICKED As Long = 0
Private Const BN_DOUBLECLICKED As Long = 5
Private Const BCN_FIRST As Long = -1250
Private Const BCN_HOTITEMCHANGE As Long = (BCN_FIRST + 1)
Private Const HICF_MOUSE As Long = &H1
Private Const HICF_ENTERING As Long = &H10
Private Const HICF_LEAVING As Long = &H20
Private Const IMAGE_BITMAP As Long = 0
Private Const IMAGE_ICON As Long = 1
Implements ISubclass
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IOleInPlaceActiveObjectVB
Implements OLEGuids.IOleControlVB
Implements OLEGuids.IPerPropertyBrowsingVB
Private CommandLinkHandle As Long
Private CommandLinkTransparentBrush As Long
Private CommandLinkAcceleratorHandle As Long
Private CommandLinkValue As Boolean
Private CommandLinkFontHandle As Long
Private CommandLinkCharCodeCache As Long
Private CommandLinkMouseOver As Boolean
Private CommandLinkDesignMode As Boolean
Private CommandLinkImageListHandle As Long
Private CommandLinkImageListObjectPointer As Long
Private UCNoSetFocusFwd As Boolean
Private DispIDMousePointer As Long
Private DispIDImageList As Long, ImageListArray() As String
Private PropDisplayAsDefault As Boolean
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropImageListName As String, PropImageListInit As Boolean
Private PropCaption As String
Private PropHint As String
Private PropPicture As IPictureDisp
Private PropTransparent As Boolean

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

Private Sub IOleControlVB_GetControlInfo(ByRef Handled As Boolean, ByRef AccelCount As Integer, ByRef AccelTable As Long, ByRef Flags As Long)
If CommandLinkAcceleratorHandle <> 0 Then
    DestroyAcceleratorTable CommandLinkAcceleratorHandle
    CommandLinkAcceleratorHandle = 0
End If
If CommandLinkHandle <> 0 Then
    Dim Accel As Integer, AccelArray() As TACCEL, AccelRefCount As Long
    Accel = AccelCharCode(Me.Caption)
    If Accel <> 0 Then
        ReDim Preserve AccelArray(0 To AccelRefCount) As TACCEL
        With AccelArray(AccelRefCount)
        .FVirt = FVIRTKEY Or FALT
        .Cmd = 1
        .Key = (VkKeyScan(Accel) And &HFF&)
        End With
        AccelRefCount = AccelRefCount + 1
        ReDim Preserve AccelArray(0 To AccelRefCount) As TACCEL
        With AccelArray(AccelRefCount)
        .FVirt = FVIRTKEY Or FALT Or FSHIFT
        .Cmd = AccelArray(AccelRefCount - 1).Cmd
        .Key = AccelArray(AccelRefCount - 1).Key
        End With
        AccelRefCount = AccelRefCount + 1
    End If
    If AccelRefCount > 0 Then
        AccelCount = AccelRefCount
        CommandLinkAcceleratorHandle = CreateAcceleratorTable(VarPtr(AccelArray(0)), AccelCount)
        AccelTable = CommandLinkAcceleratorHandle
        Flags = 0
        Handled = True
    End If
End If
End Sub

Private Sub IOleControlVB_OnMnemonic(ByRef Handled As Boolean, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Shift As Long)
If CommandLinkHandle <> 0 And wMsg = WM_SYSKEYDOWN Then
    Dim Accel As Long
    Accel = AccelCharCode(Me.Caption)
    If (VkKeyScan(Accel) And &HFF&) = (wParam And &HFF&) Then
        CommandLinkValue = True
        RaiseEvent Click
        CommandLinkValue = False
        Handled = True
    End If
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
Call SetVTableHandling(Me, VTableInterfaceControl)
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
ReDim ImageListArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
CommandLinkDesignMode = Not Ambient.UserMode
On Error GoTo 0
PropDisplayAsDefault = False
Set PropFont = Ambient.Font
PropVisualStyles = True
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftLayout = False
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
PropImageListName = "(None)"
PropCaption = Ambient.DisplayName
PropHint = vbNullString
Set PropPicture = Nothing
PropTransparent = False
Call CreateCommandLink
If CommandLinkHandle = 0 And ComCtlsSupportLevel() <= 1 And CommandLinkDesignMode = True Then
    MsgBox "The CommandLink control requires at least version 6.1 of comctl32.dll." & vbLf & _
    "In order to use it, you have to define a manifest file for your application." & vbLf & _
    "For using the control in the VB6 IDE, define a manifest file for VB6.EXE.", vbCritical + vbOKOnly
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
If DispIDImageList = 0 Then DispIDImageList = GetDispID(Me, "ImageList")
On Error Resume Next
CommandLinkDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
PropDisplayAsDefault = .ReadProperty("Default", False)
Set PropFont = .ReadProperty("Font", Nothing)
PropVisualStyles = .ReadProperty("VisualStyles", True)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
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
PropCaption = VarToStr(.ReadProperty("Caption", vbNullString))
PropHint = VarToStr(.ReadProperty("Hint", vbNullString))
Set PropPicture = .ReadProperty("Picture", Nothing)
PropTransparent = .ReadProperty("Transparent", False)
End With
Call CreateCommandLink
If Not PropImageListName = "(None)" Then TimerImageList.Enabled = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Default", PropDisplayAsDefault, False
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "VisualStyles", PropVisualStyles, True
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "ImageList", PropImageListName, "(None)"
.WriteProperty "Caption", StrToVar(PropCaption), vbNullString
.WriteProperty "Hint", StrToVar(PropHint), vbNullString
.WriteProperty "Picture", PropPicture, Nothing
.WriteProperty "Transparent", PropTransparent, False
End With
End Sub

Private Sub UserControl_Paint()
If CommandLinkHandle = 0 Then
    Dim i As Long
    For i = 8 To (UserControl.ScaleHeight + UserControl.ScaleWidth) Step 8
        UserControl.Line (-1, i)-(i, -1), vbBlack
    Next i
End If
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

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyReturn, vbKeyEscape
        CommandLinkValue = True
        RaiseEvent Click
        CommandLinkValue = False
End Select
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
Select Case PropertyName
    Case "DisplayAsDefault"
        PropDisplayAsDefault = Ambient.DisplayAsDefault
        Dim dwStyle As Long
        dwStyle = GetWindowLong(CommandLinkHandle, GWL_STYLE)
        If PropDisplayAsDefault = True Then
            If Not (dwStyle And BS_DEFPUSHBUTTON) = BS_DEFPUSHBUTTON Then
                SetWindowLong CommandLinkHandle, GWL_STYLE, dwStyle Or BS_DEFPUSHBUTTON
                Me.Refresh
            End If
        Else
            If (dwStyle And BS_DEFPUSHBUTTON) = BS_DEFPUSHBUTTON Then
                SetWindowLong CommandLinkHandle, GWL_STYLE, dwStyle And Not BS_DEFPUSHBUTTON
                Me.Refresh
            End If
        End If
End Select
End Sub

Private Sub UserControl_Resize()
Static InProc As Boolean
If InProc = True Then Exit Sub
InProc = True
With UserControl
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
If CommandLinkHandle <> 0 Then
    If PropTransparent = True Then
        MoveWindow CommandLinkHandle, 0, 0, .ScaleWidth, .ScaleHeight, 0
        If CommandLinkTransparentBrush <> 0 Then
            DeleteObject CommandLinkTransparentBrush
            CommandLinkTransparentBrush = 0
        End If
        RedrawWindow CommandLinkHandle, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
    Else
        MoveWindow CommandLinkHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    End If
End If
End With
InProc = False
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfaceInPlaceActiveObject)
Call RemoveVTableHandling(Me, VTableInterfaceControl)
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call DestroyCommandLink
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerImageList_Timer()
If PropImageListInit = False Then
    Me.ImageList = PropImageListName
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
hWnd = CommandLinkHandle
End Property

Public Property Get hWndUserControl() As Long
Attribute hWndUserControl.VB_Description = "Returns a handle to a control."
hWndUserControl = UserControl.hWnd
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object. However, this font is ignored as the control uses the current visual style system font."
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
OldFontHandle = CommandLinkFontHandle
CommandLinkFontHandle = CreateGDIFontFromOLEFont(PropFont)
If CommandLinkHandle <> 0 Then SendMessage CommandLinkHandle, WM_SETFONT, CommandLinkFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = CommandLinkFontHandle
CommandLinkFontHandle = CreateGDIFontFromOLEFont(PropFont)
If CommandLinkHandle <> 0 Then SendMessage CommandLinkHandle, WM_SETFONT, CommandLinkFontHandle, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub

Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the visual styles are enabled or not."
VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
PropVisualStyles = Value
If CommandLinkHandle <> 0 And EnabledVisualStyles() = True Then
    If PropVisualStyles = True Then
        ActivateVisualStyles CommandLinkHandle
    Else
        RemoveVisualStyles CommandLinkHandle
    End If
    Me.Refresh
End If
UserControl.PropertyChanged "VisualStyles"
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

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
If CommandLinkHandle <> 0 Then EnableWindow CommandLinkHandle, IIf(Value = True, 1, 0)
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
If CommandLinkDesignMode = False Then Call RefreshMousePointer
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
        If CommandLinkDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
If CommandLinkDesignMode = False Then Call RefreshMousePointer
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
If CommandLinkDesignMode = False Then
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
    dwMask = 0
End If
If PropRightToLeft = True Then
    If PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL Else dwMask = WS_EX_RTLREADING
End If
If CommandLinkHandle <> 0 Then Call ComCtlsSetRightToLeft(CommandLinkHandle, dwMask)
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
Attribute ImageList.VB_Description = "Returns/sets the image list control to be used. The image list should contain either a single image to be used for all states or individual images for each state."
If CommandLinkDesignMode = False Then
    If PropImageListInit = False And CommandLinkImageListObjectPointer = 0 Then
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
If CommandLinkHandle <> 0 Then
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
            CommandLinkImageListObjectPointer = ObjPtr(Value)
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
                        If CommandLinkDesignMode = False Then CommandLinkImageListObjectPointer = ObjPtr(ControlEnum)
                        PropImageListName = Value
                        Exit For
                    ElseIf CommandLinkDesignMode = True Then
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
        CommandLinkImageListObjectPointer = 0
        PropImageListName = "(None)"
    ElseIf Handle = 0 Then
        Call SetImageList(BCCL_NOGLYPH)
    End If
End If
UserControl.PropertyChanged "ImageList"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = "PPCommandLinkGeneral"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
If CommandLinkHandle <> 0 Then
    Caption = String(SendMessage(CommandLinkHandle, WM_GETTEXTLENGTH, 0, ByVal 0&), vbNullChar)
    SendMessage CommandLinkHandle, WM_GETTEXT, Len(Caption) + 1, ByVal StrPtr(Caption)
Else
    Caption = PropCaption
End If
End Property

Public Property Let Caption(ByVal Value As String)
PropCaption = Value
If CommandLinkHandle <> 0 Then
    SendMessage CommandLinkHandle, WM_SETTEXT, 0, ByVal StrPtr(PropCaption)
    Call OnControlInfoChanged(Me)
End If
UserControl.PropertyChanged "Caption"
End Property

Public Property Get Hint() As String
Attribute Hint.VB_Description = "Returns/sets the text displayed as a hint below the caption."
Attribute Hint.VB_ProcData.VB_Invoke_Property = "PPCommandLinkGeneral"
Attribute Hint.VB_UserMemId = -517
If CommandLinkHandle <> 0 Then
    Dim Length As Long
    Hint = String(SendMessage(CommandLinkHandle, BCM_GETNOTELENGTH, 0, ByVal 0&), vbNullChar)
    Length = Len(Hint) + 1 ' wParam [in, out] ; Thus the value must be stored in a variable and pointed to it.
    SendMessage CommandLinkHandle, BCM_GETNOTE, VarPtr(Length), ByVal StrPtr(Hint)
Else
    Hint = PropHint
End If
End Property

Public Property Let Hint(ByVal Value As String)
PropHint = Value
If CommandLinkHandle <> 0 Then SendMessage CommandLinkHandle, BCM_SETNOTE, 0, ByVal StrPtr(PropHint)
UserControl.PropertyChanged "Hint"
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
    If CommandLinkHandle <> 0 And CommandLinkImageListHandle = 0 Then
        dwStyle = GetWindowLong(CommandLinkHandle, GWL_STYLE)
        If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
        If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
        SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
        SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
        SetWindowLong CommandLinkHandle, GWL_STYLE, dwStyle
        Me.Refresh
    End If
Else
    Set UserControl.Picture = Value
    Set PropPicture = UserControl.Picture
    Set UserControl.Picture = Nothing
    If CommandLinkHandle <> 0 And CommandLinkImageListHandle = 0 Then
        dwStyle = GetWindowLong(CommandLinkHandle, GWL_STYLE)
        If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
        If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
        If PropPicture.Handle <> 0 Then
            If PropPicture.Type = vbPicTypeIcon Then
                dwStyle = dwStyle Or BS_ICON
                SetWindowLong CommandLinkHandle, GWL_STYLE, dwStyle
                SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
                SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_ICON, ByVal PropPicture.Handle
            Else
                dwStyle = dwStyle Or BS_BITMAP
                SetWindowLong CommandLinkHandle, GWL_STYLE, dwStyle
                SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
                SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal PropPicture.Handle
            End If
        Else
            SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
            SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
            SetWindowLong CommandLinkHandle, GWL_STYLE, dwStyle
        End If
        Me.Refresh
    End If
End If
UserControl.PropertyChanged "Picture"
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

Private Sub CreateCommandLink()
If CommandLinkHandle <> 0 Or ComCtlsSupportLevel() <= 1 Then Exit Sub
Dim dwStyle As Long, dwExStyle As Long
dwStyle = WS_CHILD Or WS_VISIBLE Or BS_COMMANDLINK Or BS_PUSHBUTTON Or BS_TEXT Or BS_NOTIFY
If PropDisplayAsDefault = True Then dwStyle = dwStyle Or BS_DEFPUSHBUTTON
If PropRightToLeft = True Then
    If PropRightToLeftLayout = True Then
        dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    Else
        dwExStyle = dwExStyle Or WS_EX_RTLREADING
    End If
End If
CommandLinkHandle = CreateWindowEx(dwExStyle, StrPtr("Button"), 0, dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
If CommandLinkHandle <> 0 Then Call ComCtlsShowAllUIStates(CommandLinkHandle)
Set Me.Font = PropFont
Me.VisualStyles = PropVisualStyles
Me.Enabled = UserControl.Enabled
Me.Caption = PropCaption
Me.Hint = PropHint
If Not PropPicture Is Nothing Then Set Me.Picture = PropPicture
If CommandLinkDesignMode = False Then
    If CommandLinkHandle <> 0 Then Call ComCtlsSetSubclass(CommandLinkHandle, Me, 1)
    Call ComCtlsSetSubclass(UserControl.hWnd, Me, 2)
End If
End Sub

Private Sub DestroyCommandLink()
If CommandLinkHandle = 0 Then Exit Sub
Call ComCtlsRemoveSubclass(CommandLinkHandle)
Call ComCtlsRemoveSubclass(UserControl.hWnd)
ShowWindow CommandLinkHandle, SW_HIDE
SetParent CommandLinkHandle, 0
DestroyWindow CommandLinkHandle
CommandLinkHandle = 0
If CommandLinkFontHandle <> 0 Then
    DeleteObject CommandLinkFontHandle
    CommandLinkFontHandle = 0
End If
If CommandLinkAcceleratorHandle <> 0 Then
    DestroyAcceleratorTable CommandLinkAcceleratorHandle
    CommandLinkAcceleratorHandle = 0
End If
If CommandLinkTransparentBrush <> 0 Then
    DeleteObject CommandLinkTransparentBrush
    CommandLinkTransparentBrush = 0
End If
CommandLinkImageListHandle = 0
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
If CommandLinkTransparentBrush <> 0 Then
    DeleteObject CommandLinkTransparentBrush
    CommandLinkTransparentBrush = 0
End If
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the value of an object."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "400"
Value = CommandLinkValue
End Property

Public Property Let Value(ByVal NewValue As Boolean)
If NewValue = True And CommandLinkValue = False Then
    CommandLinkValue = True
    RaiseEvent Click
    CommandLinkValue = False
End If
End Property

Public Sub PerformClick()
Attribute PerformClick.VB_Description = "Method that simulates a user button click."
If CommandLinkHandle <> 0 Then
    Dim hWnd As Long
    hWnd = GetAncestor(CommandLinkHandle, GA_ROOT)
    If hWnd <> 0 Then SetActiveWindow hWnd
    SetFocusAPI UserControl.hWnd
    SendMessage CommandLinkHandle, BM_CLICK, 0, ByVal 0&
End If
End Sub

Public Function SetShield(ByVal State As Boolean) As Long
Attribute SetShield.VB_Description = "Sets the elevation required state to display an elevated icon. Returns 1 if successful, or an error code otherwise."
If CommandLinkHandle <> 0 Then
    If State = True Then
        SetShield = SendMessage(CommandLinkHandle, BCM_SETSHIELD, 0, ByVal 1&)
    Else
        SetShield = SendMessage(CommandLinkHandle, BCM_SETSHIELD, 0, ByVal 0&)
        Set Me.Picture = PropPicture
    End If
End If
End Function

Public Property Get Pushed() As Boolean
Attribute Pushed.VB_Description = "Returns/sets a value that indicates if the command link is in the pushed state."
Attribute Pushed.VB_MemberFlags = "400"
If CommandLinkHandle <> 0 Then Pushed = CBool((SendMessage(CommandLinkHandle, BM_GETSTATE, 0, ByVal 0&) And BST_PUSHED) = BST_PUSHED)
End Property

Public Property Let Pushed(ByVal Value As Boolean)
If CommandLinkHandle <> 0 Then SendMessage CommandLinkHandle, BM_SETSTATE, IIf(Value = True, 1, 0), ByVal 0&
End Property

Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Returns/sets a value that indicates if the command button is hot; that is, the mouse is hovering over it. Requires comctl32.dll version 6.0 or higher."
Attribute Hot.VB_MemberFlags = "400"
If CommandLinkHandle <> 0 Then Hot = CBool((SendMessage(CommandLinkHandle, BM_GETSTATE, 0, ByVal 0&) And BST_HOT) = BST_HOT)
End Property

Public Property Let Hot(ByVal Value As Boolean)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Function GetIdealHeight() As Single
Attribute GetIdealHeight.VB_Description = "Gets the ideal height of the control."
If CommandLinkHandle <> 0 Then
    Dim Size As SIZEAPI
    SendMessage CommandLinkHandle, BCM_GETIDEALSIZE, 0, ByVal VarPtr(Size)
    ' Size.CX is not supported.
    GetIdealHeight = UserControl.ScaleY(Size.CY, vbPixels, vbContainerSize)
End If
End Function

Private Sub SetImageList(ByVal hImageList As Long)
If CommandLinkHandle <> 0 Then
    Dim BTNIML As BUTTON_IMAGELIST
    With BTNIML
    .hImageList = hImageList
    If .hImageList = 0 Then .hImageList = BCCL_NOGLYPH
    CommandLinkImageListHandle = hImageList
    If CommandLinkImageListHandle = BCCL_NOGLYPH Then CommandLinkImageListHandle = 0
    If .hImageList <> BCCL_NOGLYPH Then
        Dim dwStyle As Long
        dwStyle = GetWindowLong(CommandLinkHandle, GWL_STYLE)
        If (dwStyle And BS_ICON) = BS_ICON Then dwStyle = dwStyle And Not BS_ICON
        If (dwStyle And BS_BITMAP) = BS_BITMAP Then dwStyle = dwStyle And Not BS_BITMAP
        SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
        SendMessage CommandLinkHandle, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
        SetWindowLong CommandLinkHandle, GWL_STYLE, dwStyle
    End If
    ' .RCMargin is not supported.
    ' .uAlign is not supported.
    SendMessage CommandLinkHandle, BCM_SETIMAGELIST, 0, ByVal VarPtr(BTNIML)
    If .hImageList = BCCL_NOGLYPH Then Set Me.Picture = PropPicture
    End With
    Me.Refresh
End If
End Sub

Private Function PropImageListControl() As Object
If CommandLinkImageListObjectPointer <> 0 Then Set PropImageListControl = PtrToObj(CommandLinkImageListObjectPointer)
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
            CommandLinkCharCodeCache = ComCtlsPeekCharCode(hWnd)
        ElseIf wMsg = WM_SYSKEYDOWN Then
            RaiseEvent KeyDown(KeyCode, GetShiftStateFromMsg())
        ElseIf wMsg = WM_SYSKEYUP Then
            RaiseEvent KeyUp(KeyCode, GetShiftStateFromMsg())
        End If
        wParam = KeyCode
    Case WM_CHAR
        Dim KeyChar As Integer
        If CommandLinkCharCodeCache <> 0 Then
            KeyChar = CUIntToInt(CommandLinkCharCodeCache And &HFFFF&)
            CommandLinkCharCodeCache = 0
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
                If CommandLinkMouseOver = False And PropMouseTrack = True Then
                    CommandLinkMouseOver = True
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
        End Select
    Case WM_MOUSELEAVE
        If CommandLinkMouseOver = True Then
            CommandLinkMouseOver = False
            RaiseEvent MouseLeave
        End If
End Select
End Function

Private Function WindowProcUserControl(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case wMsg
    Case WM_COMMAND
        If lParam = CommandLinkHandle Then
            Select Case HiWord(wParam)
                Case BN_CLICKED, BN_DOUBLECLICKED
                    CommandLinkValue = True
                    RaiseEvent Click
                    CommandLinkValue = False
            End Select
        End If
    Case WM_NOTIFY
        Dim NM As NMHDR
        CopyMemory NM, ByVal lParam, LenB(NM)
        If NM.hWndFrom = CommandLinkHandle Then
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
            If CommandLinkTransparentBrush = 0 Then
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
                        CommandLinkTransparentBrush = CreatePatternBrush(hBmp)
                        SelectObject hDCBmp, hBmpOld
                        DeleteObject hBmp
                    End If
                    DeleteDC hDCBmp
                End If
            End If
            End With
            If CommandLinkTransparentBrush <> 0 Then WindowProcUserControl = CommandLinkTransparentBrush
        End If
        Exit Function
End Select
WindowProcUserControl = ComCtlsDefaultProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_SETFOCUS And UCNoSetFocusFwd = False Then SetFocusAPI CommandLinkHandle
End Function
