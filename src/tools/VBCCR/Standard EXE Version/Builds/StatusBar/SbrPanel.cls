VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SbrPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShadowObjectPointer As Long
Private PropIndex As Long, PropKey As String
Private PropTag As Variant

Private Function ShadowStatusBar() As StatusBar
ComCtlsPtrToShadowObj ShadowStatusBar, ShadowObjectPointer
End Function

Friend Sub FInit(ByVal Object As StatusBar, ByVal Key As String)
ShadowObjectPointer = ObjPtr(Object)
PropKey = Key
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the panel."
Index = PropIndex
End Property

Friend Property Let Index(ByVal Value As Long)
PropIndex = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = PropKey
End Property

Public Property Let Key(ByVal Value As String)
ShadowStatusBar.Panels.FChangeKey PropIndex, PropKey, Value
End Property

Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
If IsEmpty(PropTag) Then
    Tag = vbNullString
ElseIf IsObject(PropTag) Then
    Set Tag = PropTag
Else
    Tag = PropTag
End If
End Property

Public Property Let Tag(ByVal Value As Variant)
PropTag = Value
End Property

Public Property Set Tag(ByVal Value As Variant)
If Value Is ShadowStatusBar Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text of the panel."
Text = ShadowStatusBar.FPanelText(PropIndex)
End Property

Public Property Let Text(ByVal Value As String)
ShadowStatusBar.FPanelText(PropIndex) = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the panel. Requires that the show tips property is set to true."
ToolTipText = ShadowStatusBar.FPanelToolTipText(PropIndex)
End Property

Public Property Let ToolTipText(ByVal Value As String)
ShadowStatusBar.FPanelToolTipText(PropIndex) = Value
End Property

Public Property Get Style() As SbrPanelStyleConstants
Attribute Style.VB_Description = "Returns/sets the style of the panel."
Style = ShadowStatusBar.FPanelStyle(PropIndex)
End Property

Public Property Let Style(ByVal Value As SbrPanelStyleConstants)
ShadowStatusBar.FPanelStyle(PropIndex) = Value
End Property

Public Property Get Bevel() As SbrPanelBevelConstants
Attribute Bevel.VB_Description = "Returns/sets the bevel appearance of the panel."
Bevel = ShadowStatusBar.FPanelBevel(PropIndex)
End Property

Public Property Let Bevel(ByVal Value As SbrPanelBevelConstants)
ShadowStatusBar.FPanelBevel(PropIndex) = Value
End Property

Public Property Get AutoSize() As SbrPanelAutoSizeConstants
Attribute AutoSize.VB_Description = "Returns/sets a value that allows the width of the panel to be automatically sized."
AutoSize = ShadowStatusBar.FPanelAutoSize(PropIndex)
End Property

Public Property Let AutoSize(ByVal Value As SbrPanelAutoSizeConstants)
ShadowStatusBar.FPanelAutoSize(PropIndex) = Value
End Property

Public Property Get Alignment() As SbrPanelAlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of the panel."
Alignment = ShadowStatusBar.FPanelAlignment(PropIndex)
End Property

Public Property Let Alignment(ByVal Value As SbrPanelAlignmentConstants)
ShadowStatusBar.FPanelAlignment(PropIndex) = Value
End Property

Public Property Get DTFormat() As SbrPanelDTFormatConstants
Attribute DTFormat.VB_Description = "Returns/sets the date and time format of the panel."
DTFormat = ShadowStatusBar.FPanelDTFormat(PropIndex)
End Property

Public Property Let DTFormat(ByVal Value As SbrPanelDTFormatConstants)
ShadowStatusBar.FPanelDTFormat(PropIndex) = Value
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = ShadowStatusBar.FPanelForeColor(PropIndex)
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
ShadowStatusBar.FPanelForeColor(PropIndex) = Value
End Property

Public Property Get MinWidth() As Single
Attribute MinWidth.VB_Description = "Returns/sets the minimum width of the panel."
MinWidth = ShadowStatusBar.FPanelMinWidth(PropIndex)
End Property

Public Property Let MinWidth(ByVal Value As Single)
ShadowStatusBar.FPanelMinWidth(PropIndex) = Value
End Property

Public Property Get Picture() As IPictureDisp
Attribute Picture.VB_Description = "Returns/sets the picture of the panel."
Set Picture = ShadowStatusBar.FPanelPicture(PropIndex)
End Property

Public Property Let Picture(ByVal Value As IPictureDisp)
ShadowStatusBar.FPanelPicture(PropIndex) = Value
End Property

Public Property Set Picture(ByVal Value As IPictureDisp)
Set ShadowStatusBar.FPanelPicture(PropIndex) = Value
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether the panel text appears as enabled or not."
Enabled = ShadowStatusBar.FPanelEnabled(PropIndex)
End Property

Public Property Let Enabled(ByVal Value As Boolean)
ShadowStatusBar.FPanelEnabled(PropIndex) = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether the panel is visible or not."
Visible = ShadowStatusBar.FPanelVisible(PropIndex)
End Property

Public Property Let Visible(ByVal Value As Boolean)
ShadowStatusBar.FPanelVisible(PropIndex) = Value
End Property

Public Property Get Bold() As Boolean
Attribute Bold.VB_Description = "Returns/sets a value that determines whether the text is displayed with a bold font or not."
Bold = ShadowStatusBar.FPanelBold(PropIndex)
End Property

Public Property Let Bold(ByVal Value As Boolean)
ShadowStatusBar.FPanelBold(PropIndex) = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container."
Left = ShadowStatusBar.FPanelLeft(PropIndex)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns/sets the width of the panel."
Width = ShadowStatusBar.FPanelWidth(PropIndex)
End Property

Public Property Let Width(ByVal Value As Single)
ShadowStatusBar.FPanelWidth(PropIndex) = Value
End Property
