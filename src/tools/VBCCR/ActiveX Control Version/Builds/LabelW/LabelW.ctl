VERSION 5.00
Begin VB.UserControl LabelW 
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DataBindingBehavior=   1  'vbSimpleBound
   DrawStyle       =   5  'Transparent
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   PropertyPages   =   "LabelW.ctx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ToolboxBitmap   =   "LabelW.ctx":0035
   Windowless      =   -1  'True
   Begin VB.Timer TimerMouseTrack 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "LabelW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private LblEllipsisFormatNone, LblEllipsisFormatEnd, LblEllipsisFormatPath, LblEllipsisFormatWord
#End If
Public Enum LblEllipsisFormatConstants
LblEllipsisFormatNone = 0
LblEllipsisFormatEnd = 1
LblEllipsisFormatPath = 2
LblEllipsisFormatWord = 3
End Enum
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type TMSG
hWnd As Long
Message As Long
wParam As Long
lParam As Long
Time As Long
PT As POINTAPI
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
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
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByRef lpMsg As TMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (ByRef lpMsg As TMSG) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function GetMessagePos Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, ByRef qRC As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Const DT_LEFT As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_RIGHT As Long = &H2
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_NOCLIP As Long = &H100
Private Const DT_CALCRECT As Long = &H400
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_PATH_ELLIPSIS As Long = &H4000
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_RTLREADING As Long = &H20000
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const SM_CXBORDER As Long = 5
Private Const SM_CYBORDER As Long = 6
Private Const SM_CXEDGE As Long = 45
Private Const SM_CYEDGE As Long = 46
Private Const SM_CXDLGFRAME As Long = 7
Private Const SM_CYDLGFRAME As Long = 8
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKENINNER As Long = &H8
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BF_LEFT As Long = &H1
Private Const BF_RIGHT As Long = &H4
Private Const BF_TOP As Long = &H2
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Implements OLEGuids.IObjectSafety
Implements OLEGuids.IPerPropertyBrowsingVB
Private LabelAutoSizeFlag As Boolean
Private LabelDisplayedCaption As String
Private LabelPaintFrozen As Boolean
Private LabelMouseOver As Boolean, LabelMouseOverPos As Long
Private LabelDesignMode As Boolean
Private DispIDMousePointer As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropMouseTrack As Boolean
Private PropRightToLeft As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropAlignment As VBRUN.AlignmentConstants
Private PropBorderStyle As CCBorderStyleConstants
Private PropCaption As String
Private PropUseMnemonic As Boolean
Private PropAutoSize As Boolean
Private PropWordWrap As Boolean
Private PropEllipsisFormat As LblEllipsisFormatConstants
Private PropVerticalAlignment As CCVerticalAlignmentConstants

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
Call SetVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
End Sub

Private Sub UserControl_InitProperties()
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
LabelDesignMode = Not Ambient.UserMode
On Error GoTo 0
Set PropFont = Ambient.Font
Set UserControl.Font = PropFont
PropMousePointer = 0: Set PropMouseIcon = Nothing
PropMouseTrack = False
PropRightToLeft = Ambient.RightToLeft
PropRightToLeftMode = CCRightToLeftModeVBAME
If PropRightToLeft = True Then Me.RightToLeft = True
If PropRightToLeft = False Then PropAlignment = vbLeftJustify Else PropAlignment = vbRightJustify
PropBorderStyle = CCBorderStyleNone
PropCaption = Ambient.DisplayName
PropUseMnemonic = True
PropAutoSize = False
PropWordWrap = False
PropEllipsisFormat = LblEllipsisFormatNone
PropVerticalAlignment = CCVerticalAlignmentTop
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
If DispIDMousePointer = 0 Then DispIDMousePointer = GetDispID(Me, "MousePointer")
On Error Resume Next
LabelDesignMode = Not Ambient.UserMode
On Error GoTo 0
With PropBag
Set PropFont = .ReadProperty("Font", Nothing)
If PropFont Is Nothing Then Set PropFont = Ambient.Font
Set UserControl.Font = PropFont
Me.Appearance = .ReadProperty("Appearance", CCAppearance3D)
Me.BackColor = .ReadProperty("BackColor", vbButtonFace)
Me.ForeColor = .ReadProperty("ForeColor", vbButtonText)
Me.Enabled = .ReadProperty("Enabled", True)
Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
PropMousePointer = .ReadProperty("MousePointer", 0)
Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
Me.MousePointer = PropMousePointer
PropMouseTrack = .ReadProperty("MouseTrack", False)
PropRightToLeft = .ReadProperty("RightToLeft", False)
PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
If PropRightToLeft = True Then Me.RightToLeft = True
PropAlignment = .ReadProperty("Alignment", vbLeftJustify)
PropBorderStyle = .ReadProperty("BorderStyle", CCBorderStyleNone)
If PropBorderStyle = CCBorderStyleSingle Then UserControl.DrawStyle = vbSolid Else UserControl.DrawStyle = vbInvisible
Me.BackStyle = .ReadProperty("BackStyle", CCBackStyleOpaque)
PropCaption = .ReadProperty("Caption", vbNullString) ' Unicode not necessary
PropUseMnemonic = .ReadProperty("UseMnemonic", True)
PropAutoSize = .ReadProperty("AutoSize", False)
PropWordWrap = .ReadProperty("WordWrap", False)
PropEllipsisFormat = .ReadProperty("EllipsisFormat", LblEllipsisFormatNone)
PropVerticalAlignment = .ReadProperty("VerticalAlignment", CCVerticalAlignmentTop)
End With
If PropUseMnemonic = True Then
    UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
Else
    UserControl.AccessKeys = vbNullString
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
.WriteProperty "Appearance", Me.Appearance, CCAppearance3D
.WriteProperty "BackColor", Me.BackColor, vbButtonFace
.WriteProperty "ForeColor", Me.ForeColor, vbButtonText
.WriteProperty "Enabled", Me.Enabled, True
.WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
.WriteProperty "MousePointer", PropMousePointer, 0
.WriteProperty "MouseIcon", PropMouseIcon, Nothing
.WriteProperty "MouseTrack", PropMouseTrack, False
.WriteProperty "RightToLeft", PropRightToLeft, False
.WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
.WriteProperty "Alignment", PropAlignment, vbLeftJustify
.WriteProperty "BorderStyle", PropBorderStyle, CCBorderStyleNone
.WriteProperty "BackStyle", Me.BackStyle, CCBackStyleOpaque
.WriteProperty "Caption", PropCaption, vbNullString ' Unicode not necessary
.WriteProperty "UseMnemonic", PropUseMnemonic, True
.WriteProperty "AutoSize", PropAutoSize, False
.WriteProperty "WordWrap", PropWordWrap, False
.WriteProperty "EllipsisFormat", PropEllipsisFormat, LblEllipsisFormatNone
.WriteProperty "VerticalAlignment", PropVerticalAlignment, CCVerticalAlignmentTop
End With
End Sub

Private Sub UserControl_Paint()
If LabelPaintFrozen = True Then Exit Sub
Dim RC As RECT, CalcRect As RECT, Format As Long, Buffer As String
Dim BorderWidth As Long, BorderHeight As Long
With UserControl
SetRect RC, 0, 0, .ScaleWidth, .ScaleHeight
Select Case PropBorderStyle
    Case CCBorderStyleSingle
        BorderWidth = GetSystemMetrics(SM_CXBORDER)
        BorderHeight = GetSystemMetrics(SM_CYBORDER)
        UserControl.Line (0, 0)-(.ScaleWidth - BorderWidth, .ScaleHeight - BorderHeight), RGB(100, 100, 100), B
    Case CCBorderStyleThin
        BorderWidth = GetSystemMetrics(SM_CXBORDER)
        BorderHeight = GetSystemMetrics(SM_CYBORDER)
        DrawEdge .hDC, RC, BDR_SUNKENOUTER, BF_RECT
    Case CCBorderStyleSunken
        BorderWidth = GetSystemMetrics(SM_CXEDGE)
        BorderHeight = GetSystemMetrics(SM_CYEDGE)
        DrawEdge .hDC, RC, BDR_SUNKENOUTER Or BDR_SUNKENINNER, BF_RECT
    Case CCBorderStyleRaised
        BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
        BorderHeight = GetSystemMetrics(SM_CYDLGFRAME)
        DrawEdge .hDC, RC, BDR_RAISEDOUTER Or BDR_RAISEDINNER, BF_RECT
End Select
If .Enabled = False Then SetTextColor .hDC, WinColor(vbGrayText)
Format = DT_NOCLIP
Select Case PropAlignment
    Case vbLeftJustify
        Format = Format Or DT_LEFT
    Case vbCenter
        Format = Format Or DT_CENTER
    Case vbRightJustify
        Format = Format Or DT_RIGHT
End Select
If PropRightToLeft = True Then Format = Format Or DT_RTLREADING
If PropUseMnemonic = False Then Format = Format Or DT_NOPREFIX
If PropWordWrap = True Then Format = Format Or DT_WORDBREAK
Select Case PropEllipsisFormat
    Case LblEllipsisFormatEnd
        Format = Format Or DT_END_ELLIPSIS
    Case LblEllipsisFormatPath
        Format = Format Or DT_PATH_ELLIPSIS
    Case LblEllipsisFormatWord
        Format = Format Or DT_WORD_ELLIPSIS
End Select
If PropVerticalAlignment <> CCVerticalAlignmentTop Then
    Dim Height As Long, Result As Long
    Buffer = PropCaption
    If Buffer = vbNullString Then Buffer = " "
    LSet CalcRect = RC
    Height = DrawText(.hDC, StrPtr(Buffer), -1, CalcRect, Format Or DT_CALCRECT)
    Select Case PropVerticalAlignment
        Case CCVerticalAlignmentCenter
            Result = ((((RC.Bottom - RC.Top) - (BorderHeight * 2)) - Height) / 2)
        Case CCVerticalAlignmentBottom
            Result = (((RC.Bottom - RC.Top) - (BorderHeight * 2)) - Height)
    End Select
    If Result > 0 Then RC.Top = RC.Top + Result
End If
SetRect RC, RC.Left + BorderWidth, RC.Top + BorderHeight, RC.Right - (BorderWidth * 2), RC.Bottom - (BorderHeight * 2)
If Not PropCaption = vbNullString Then
    ' The function could add up to four additional characters to this string.
    ' The buffer containing the string should be large enough to accommodate these extra characters.
    Buffer = PropCaption & String$(4, vbNullChar) & vbNullChar
    Dim hRgn As Long
    If PropAutoSize = True Then
        ' Temporarily remove the clipping region in this case.
        hRgn = CreateRectRgn(0, 0, 0, 0)
        If hRgn <> 0 Then
            If GetClipRgn(.hDC, hRgn) = 1 Then
                SelectClipRgn .hDC, 0
            Else
                DeleteObject hRgn
                hRgn = 0
            End If
        End If
    End If
    DrawText .hDC, StrPtr(Buffer), -1, RC, Format Or DT_MODIFYSTRING
    If hRgn <> 0 Then
        SelectClipRgn .hDC, hRgn
        DeleteObject hRgn
        hRgn = 0
    End If
    LabelDisplayedCaption = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
Else
    LabelDisplayedCaption = vbNullString
End If
End With
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, UserControl.ScaleX(X, vbPixels, vbTwips), UserControl.ScaleY(Y, vbPixels, vbTwips))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If LabelMouseOver = False And PropMouseTrack = True Then
    LabelMouseOver = True
    RaiseEvent MouseEnter
    LabelMouseOverPos = GetMessagePos()
    TimerMouseTrack.Enabled = True
ElseIf LabelMouseOver = True Then
    LabelMouseOverPos = GetMessagePos()
End If
RaiseEvent MouseMove(Button, Shift, UserControl.ScaleX(X, vbPixels, vbTwips), UserControl.ScaleY(Y, vbPixels, vbTwips))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, UserControl.ScaleX(X, vbPixels, vbTwips), UserControl.ScaleY(Y, vbPixels, vbTwips))
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
If DPICorrectionFactor() <> 1 Then Call SyncObjectRectsToContainer(Me)
Call RedrawLabel
InProc = False
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
If HitResult = vbHitResultOutside Then HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Terminate()
Call RemoveVTableHandling(Me, VTableInterfacePerPropertyBrowsing)
Call ComCtlsReleaseShellMod
End Sub

Private Sub TimerMouseTrack_Timer()
If GetCapture() <> UserControl.ContainerHwnd Then
    Dim Pos As Long
    Pos = GetMessagePos()
    If LabelMouseOverPos <> Pos Or WindowFromPoint(Get_X_lParam(Pos), Get_Y_lParam(Pos)) <> UserControl.ContainerHwnd Then
        LabelMouseOver = False
        TimerMouseTrack.Enabled = False
        RaiseEvent MouseLeave
    End If
Else
    LabelMouseOverPos = 0
End If
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
Set PropFont = NewFont
Set UserControl.Font = PropFont
LabelAutoSizeFlag = PropAutoSize
Call RedrawLabel
UserControl.PropertyChanged "Font"
End Property

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Set UserControl.Font = PropFont
LabelAutoSizeFlag = PropAutoSize
Call RedrawLabel
UserControl.PropertyChanged "Font"
End Sub

Public Property Get Appearance() As CCAppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets a value that determines whether an object is painted two-dimensional or with 3-D effects."
Attribute Appearance.VB_UserMemId = -520
Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal Value As CCAppearanceConstants)
Select Case Value
    Case CCAppearanceFlat, CCAppearance3D
        UserControl.Appearance = Value
    Case Else
        Err.Raise 380
End Select
UserControl.ForeColor = IIf(UserControl.Appearance = CCAppearanceFlat, vbWindowText, vbButtonText)
If UserControl.Appearance = CCAppearanceFlat Then
    If Not PropBorderStyle = CCBorderStyleNone Then PropBorderStyle = CCBorderStyleSingle
Else
    If Not PropBorderStyle = CCBorderStyleNone Then PropBorderStyle = CCBorderStyleSunken
End If
Call RedrawLabel
UserControl.PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
UserControl.BackColor = Value
Call RedrawLabel
UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
UserControl.ForeColor = Value
Call RedrawLabel
UserControl.PropertyChanged "ForeColor"
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
If LabelDesignMode = False Then
    Select Case PropMousePointer
        Case vbIconPointer, 16, vbCustom
            If PropMousePointer = vbCustom Then
                Set UserControl.MouseIcon = PropMouseIcon
            Else
                Set UserControl.MouseIcon = PictureFromHandle(LoadCursor(0, MousePointerID(PropMousePointer)), vbPicTypeIcon)
            End If
            UserControl.MousePointer = vbCustom
        Case Else
            UserControl.MousePointer = PropMousePointer
    End Select
End If
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
        If LabelDesignMode = True Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
End If
Me.MousePointer = PropMousePointer
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
Call RedrawLabel
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

Public Property Get Alignment() As VBRUN.AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment."
Alignment = PropAlignment
End Property

Public Property Let Alignment(ByVal Value As VBRUN.AlignmentConstants)
Select Case Value
    Case vbLeftJustify, vbCenter, vbRightJustify
        PropAlignment = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawLabel
UserControl.PropertyChanged "TextAlignment"
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
        If PropBorderStyle = CCBorderStyleSingle Then UserControl.DrawStyle = vbSolid Else UserControl.DrawStyle = vbInvisible
    Case Else
        Err.Raise 380
End Select
LabelAutoSizeFlag = PropAutoSize
Call RedrawLabel
UserControl.PropertyChanged "BorderStyle"
End Property

Public Property Get BackStyle() As CCBackStyleConstants
Attribute BackStyle.VB_Description = "Returns/sets the background style."
Attribute BackStyle.VB_UserMemId = -502
BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal Value As CCBackStyleConstants)
Select Case Value
    Case CCBackStyleTransparent, CCBackStyleOpaque
        UserControl.BackStyle = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawLabel
UserControl.PropertyChanged "BackStyle"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "123c"
Caption = PropCaption
End Property

Public Property Let Caption(ByVal Value As String)
If PropCaption = Value Then Exit Property
PropCaption = Value
If PropUseMnemonic = True Then UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
LabelAutoSizeFlag = PropAutoSize
Call RedrawLabel
UserControl.PropertyChanged "Caption"
On Error Resume Next
UserControl.Extender.DataChanged = True
On Error GoTo 0
RaiseEvent Change
End Property

Public Property Get Default() As String
Attribute Default.VB_UserMemId = 0
Attribute Default.VB_MemberFlags = "40"
Default = Me.Caption
End Property

Public Property Let Default(ByVal Value As String)
Me.Caption = Value
End Property

Public Property Get UseMnemonic() As Boolean
Attribute UseMnemonic.VB_Description = "Returns/sets a value that specifies whether an & in the caption property defines an access key."
UseMnemonic = PropUseMnemonic
End Property

Public Property Let UseMnemonic(ByVal Value As Boolean)
PropUseMnemonic = Value
If PropUseMnemonic = True Then
    UserControl.AccessKeys = ChrW(AccelCharCode(PropCaption))
Else
    UserControl.AccessKeys = vbNullString
End If
LabelAutoSizeFlag = PropAutoSize
Call RedrawLabel
UserControl.PropertyChanged "UseMnemonic"
End Property

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
Attribute AutoSize.VB_UserMemId = -500
AutoSize = PropAutoSize
End Property

Public Property Let AutoSize(ByVal Value As Boolean)
PropAutoSize = Value
LabelAutoSizeFlag = PropAutoSize
Call RedrawLabel
UserControl.PropertyChanged "AutoSize"
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control may break lines within the text in order to prevent overflow."
WordWrap = PropWordWrap
End Property

Public Property Let WordWrap(ByVal Value As Boolean)
PropWordWrap = Value
LabelAutoSizeFlag = PropAutoSize
Call RedrawLabel
UserControl.PropertyChanged "WordWrap"
End Property

Public Property Get EllipsisFormat() As LblEllipsisFormatConstants
Attribute EllipsisFormat.VB_Description = "Returns/sets a value indicating if and where the ellipsis character is appended, denoting that the caption extends beyond the length of the label. The auto size and the word wrap property may be set to false to see the ellipsis character."
EllipsisFormat = PropEllipsisFormat
End Property

Public Property Let EllipsisFormat(ByVal Value As LblEllipsisFormatConstants)
Select Case Value
    Case LblEllipsisFormatNone, LblEllipsisFormatEnd, LblEllipsisFormatPath, LblEllipsisFormatWord
        PropEllipsisFormat = Value
    Case Else
        Err.Raise 380
End Select
LabelAutoSizeFlag = PropAutoSize
Call RedrawLabel
UserControl.PropertyChanged "EllipsisFormat"
End Property

Public Property Get VerticalAlignment() As CCVerticalAlignmentConstants
Attribute VerticalAlignment.VB_Description = "Returns/sets the vertical alignment."
VerticalAlignment = PropVerticalAlignment
End Property

Public Property Let VerticalAlignment(ByVal Value As CCVerticalAlignmentConstants)
Select Case Value
    Case CCVerticalAlignmentTop, CCVerticalAlignmentCenter, CCVerticalAlignmentBottom
        PropVerticalAlignment = Value
    Case Else
        Err.Raise 380
End Select
Call RedrawLabel
UserControl.PropertyChanged "VerticalAlignment"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
Call RedrawLabel
Dim Msg As TMSG, hWndContainer As Long
hWndContainer = UserControl.ContainerHwnd
Const WM_PAINT As Long = &HF, PM_REMOVE As Long = &H1
Do While PeekMessage(Msg, hWndContainer, WM_PAINT, WM_PAINT, PM_REMOVE) <> 0
    DispatchMessage Msg
Loop
End Sub

Public Property Get DisplayedCaption() As String
Attribute DisplayedCaption.VB_Description = "Returns the modified string to match the displayed text."
Attribute DisplayedCaption.VB_MemberFlags = "400"
DisplayedCaption = LabelDisplayedCaption
End Property

Private Sub DoAutoSize(ByVal hDC As Long)
Dim RC As RECT, CalcRect As RECT, Format As Long, Buffer As String
Dim BorderWidth As Long, BorderHeight As Long
With UserControl
SetRect RC, 0, 0, .ScaleWidth, .ScaleHeight
Select Case PropBorderStyle
    Case CCBorderStyleSingle
        BorderWidth = GetSystemMetrics(SM_CXBORDER)
        BorderHeight = GetSystemMetrics(SM_CYBORDER)
    Case CCBorderStyleThin
        BorderWidth = GetSystemMetrics(SM_CXBORDER)
        BorderHeight = GetSystemMetrics(SM_CYBORDER)
    Case CCBorderStyleSunken
        BorderWidth = GetSystemMetrics(SM_CXEDGE)
        BorderHeight = GetSystemMetrics(SM_CYEDGE)
    Case CCBorderStyleRaised
        BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
        BorderHeight = GetSystemMetrics(SM_CYDLGFRAME)
End Select
Format = DT_NOCLIP
Select Case PropAlignment
    Case vbLeftJustify
        Format = Format Or DT_LEFT
    Case vbCenter
        Format = Format Or DT_CENTER
    Case vbRightJustify
        Format = Format Or DT_RIGHT
End Select
If PropRightToLeft = True Then Format = Format Or DT_RTLREADING
If PropUseMnemonic = False Then Format = Format Or DT_NOPREFIX
If PropWordWrap = True Then Format = Format Or DT_WORDBREAK
Buffer = PropCaption
If Buffer = vbNullString Then Buffer = " "
LSet CalcRect = RC
DrawText hDC, StrPtr(Buffer), -1, CalcRect, Format Or DT_CALCRECT
Dim OldRight As Single, OldCenter As Single, OldBottom As Single, OldVCenter As Single
OldRight = .Extender.Left + .Extender.Width
OldCenter = .Extender.Left + (.Extender.Width / 2)
OldBottom = .Extender.Top + .Extender.Height
OldVCenter = .Extender.Top + (.Extender.Height / 2)
If PropWordWrap = True Then
    If .ScaleWidth < ((CalcRect.Right - CalcRect.Left) + (BorderWidth * 2)) Then
        .Extender.Move .Extender.Left, .Extender.Top, .ScaleX((CalcRect.Right - CalcRect.Left) + (BorderWidth * 2), vbPixels, vbContainerSize), .ScaleY((CalcRect.Bottom - CalcRect.Top) + (BorderHeight * 2), vbPixels, vbContainerSize)
    Else
        .Extender.Height = .ScaleY((CalcRect.Bottom - CalcRect.Top) + (BorderHeight * 2), vbPixels, vbContainerSize)
    End If
Else
    .Extender.Move .Extender.Left, .Extender.Top, .ScaleX((CalcRect.Right - CalcRect.Left) + (BorderWidth * 2), vbPixels, vbContainerSize), .ScaleY((CalcRect.Bottom - CalcRect.Top) + (BorderHeight * 2), vbPixels, vbContainerSize)
End If
LabelPaintFrozen = True
Select Case PropAlignment
    Case vbCenter
        If .Extender.Left <> (OldCenter - (.Extender.Width / 2)) Then .Extender.Left = (OldCenter - (.Extender.Width / 2))
    Case vbRightJustify
        If .Extender.Left <> (OldRight - .Extender.Width) Then .Extender.Left = (OldRight - .Extender.Width)
End Select
Select Case PropVerticalAlignment
    Case CCVerticalAlignmentCenter
        If .Extender.Top <> (OldVCenter - (.Extender.Height / 2)) Then .Extender.Top = (OldVCenter - (.Extender.Height / 2))
    Case CCVerticalAlignmentBottom
        If .Extender.Top <> (OldBottom - .Extender.Height) Then .Extender.Top = (OldBottom - .Extender.Height)
End Select
LabelPaintFrozen = False
.Refresh
End With
End Sub

Private Sub RedrawLabel()
If LabelAutoSizeFlag = False Then
    UserControl.Refresh
Else
    Dim hDCScreen As Long, hDC As Long
    hDCScreen = GetDC(0)
    If hDCScreen <> 0 Then
        hDC = CreateCompatibleDC(hDCScreen)
        If hDC <> 0 Then
            Dim Font As IFont, hFontOld As Long
            Set Font = PropFont
            If Not Font Is Nothing Then hFontOld = SelectObject(hDC, Font.hFont)
            Call DoAutoSize(hDC)
            If hFontOld <> 0 Then SelectObject hDC, hFontOld
            Set Font = Nothing
            DeleteDC hDC
        End If
        ReleaseDC 0, hDCScreen
    End If
    LabelAutoSizeFlag = False
End If
End Sub
