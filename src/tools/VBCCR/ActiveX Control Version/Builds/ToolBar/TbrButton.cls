VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbrButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
Private ShadowObjectPointer As LongPtr
Private PropIndex As Long, PropKey As String, PropID As Long
Private PropTag As Variant
Private PropCaption As String
Private PropImage As Variant, PropImageIndex As Long
Private PropToolTipText As String
Private PropDescription As String
Private PropForeColor As OLE_COLOR
Private PropButtonMenus As TbrButtonMenus

Private Function ShadowToolBar() As ToolBar
ComCtlsObjSetAddRef ShadowToolBar, ShadowObjectPointer
End Function

Private Sub Class_Initialize()
PropForeColor = vbButtonText
End Sub

Friend Sub FInit(ByVal Object As ToolBar, ByVal Key As String, ByVal Caption As String, ByVal Image As Variant, ByVal ImageIndex As Long)
ShadowObjectPointer = ObjPtr(Object)
PropKey = Key
PropCaption = Caption
PropImage = Image
PropImageIndex = ImageIndex
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the button."
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
ShadowToolBar.Buttons.FChangeKey PropIndex, PropKey, Value
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
If Value Is ShadowToolBar Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
Else
    Set PropTag = Value
End If
End Property

Public Property Get ID() As Long
Attribute ID.VB_Description = "Returns an internal identification of this button."
ID = PropID
End Property

Friend Property Let ID(ByVal Value As Long)
PropID = Value
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the caption of the button."
Attribute Caption.VB_UserMemId = 0
Caption = PropCaption
End Property

Public Property Let Caption(ByVal Value As String)
PropCaption = Value
ShadowToolBar.FButtonCaption(PropID) = PropCaption
End Property

Public Property Get Style() As TbrButtonStyleConstants
Attribute Style.VB_Description = "Returns/sets the style of the button."
Style = ShadowToolBar.FButtonStyle(PropID, PropImageIndex)
End Property

Public Property Let Style(ByVal Value As TbrButtonStyleConstants)
ShadowToolBar.FButtonStyle(PropID, PropImageIndex) = Value
End Property

Public Property Get Image() As Variant
Attribute Image.VB_Description = "Returns/sets the index or key of an image in an associated image list control."
If IsEmpty(PropImage) Then
    Image = 0
Else
    Image = PropImage
End If
End Property

Public Property Let Image(ByVal Value As Variant)
Call ComCtlsImlListImageIndex(ShadowToolBar, ShadowToolBar.ImageList, Value, PropImageIndex)
PropImage = Value
ShadowToolBar.FButtonImage(PropID) = PropImageIndex
End Property

Public Property Get ImageIndex() As Long
Attribute ImageIndex.VB_Description = "Returns the index of an image in an associated image list control."
ImageIndex = PropImageIndex
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the button. Requires that the show tips property is set to true."
ToolTipText = PropToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
PropToolTipText = Value
End Property

Public Property Get Description() As String
Attribute Description.VB_Description = "Returns/sets the description text displayed during a customization operation of the button."
Description = PropDescription
End Property

Public Property Let Description(ByVal Value As String)
PropDescription = Value
End Property

Public Property Get Value() As TbrButtonValueConstants
Attribute Value.VB_Description = "Returns/sets the value of the button."
Value = ShadowToolBar.FButtonValue(PropID)
End Property

Public Property Let Value(ByVal NewValue As TbrButtonValueConstants)
ShadowToolBar.FButtonValue(PropID) = NewValue
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether the button is enabled or not."
Enabled = ShadowToolBar.FButtonEnabled(PropID)
End Property

Public Property Let Enabled(ByVal Value As Boolean)
ShadowToolBar.FButtonEnabled(PropID) = Value
End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets a value that determines whether the button menu is visible or not."
Visible = ShadowToolBar.FButtonVisible(PropID)
End Property

Public Property Let Visible(ByVal Value As Boolean)
ShadowToolBar.FButtonVisible(PropID) = Value
End Property

Public Property Get MixedState() As Boolean
Attribute MixedState.VB_Description = "Returns/sets a value that determines if a button appears in an indeterminate state."
MixedState = ShadowToolBar.FButtonMixedState(PropID)
End Property

Public Property Let MixedState(ByVal Value As Boolean)
ShadowToolBar.FButtonMixedState(PropID) = Value
End Property

Public Property Get HighLighted() As Boolean
Attribute HighLighted.VB_Description = "Returns/sets a value that determines whether the button is highlighted or not. This has no visible effect if the version of comctl32.dll is 6.0 or higher and the visual styles property is set to true."
HighLighted = ShadowToolBar.FButtonHighLighted(PropID)
End Property

Public Property Let HighLighted(ByVal Value As Boolean)
ShadowToolBar.FButtonHighLighted(PropID) = Value
End Property

Public Property Get NoImage() As Boolean
Attribute NoImage.VB_Description = "Returns/sets a value indicating that no image will be drawn."
NoImage = ShadowToolBar.FButtonNoImage(PropID, PropImageIndex)
End Property

Public Property Let NoImage(ByVal Value As Boolean)
ShadowToolBar.FButtonNoImage(PropID, PropImageIndex) = Value
End Property

Public Property Get NoPrefix() As Boolean
Attribute NoPrefix.VB_Description = "Returns/sets a value indicating if the button text will not have an accelerator prefix associated with it."
NoPrefix = ShadowToolBar.FButtonNoPrefix(PropID)
End Property

Public Property Let NoPrefix(ByVal Value As Boolean)
ShadowToolBar.FButtonNoPrefix(PropID) = Value
End Property

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Returns/sets a value that indicating if the button should not assign the standard width. Instead, the width will be calculated based on the text plus the image."
AutoSize = ShadowToolBar.FButtonAutoSize(PropID)
End Property

Public Property Let AutoSize(ByVal Value As Boolean)
ShadowToolBar.FButtonAutoSize(PropID) = Value
End Property

Public Property Get CustomWidth() As Single
Attribute CustomWidth.VB_Description = "Returns/sets the custom width of the button."
CustomWidth = ShadowToolBar.FButtonCustomWidth(PropID)
End Property

Public Property Let CustomWidth(ByVal Value As Single)
ShadowToolBar.FButtonCustomWidth(PropID) = Value
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
PropForeColor = Value
ShadowToolBar.FButtonRedraw PropID
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns/sets the position of the button."
Position = ShadowToolBar.FButtonPosition(PropID)
End Property

Public Property Let Position(ByVal Value As Long)
ShadowToolBar.FButtonPosition(PropID) = Value
End Property

Public Property Get Hot() As Boolean
Attribute Hot.VB_Description = "Returns/sets a value that determines whether the button is hot or not."
Hot = ShadowToolBar.FButtonHot(PropID)
End Property

Public Property Let Hot(ByVal Value As Boolean)
ShadowToolBar.FButtonHot(PropID) = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container."
Left = ShadowToolBar.FButtonLeft(PropID)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns the distance between the internal top edge of an object and the top edge of its container."
Top = ShadowToolBar.FButtonTop(PropID)
End Property

Public Property Let Top(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the button."
Width = ShadowToolBar.FButtonWidth(PropID)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the button."
Height = ShadowToolBar.FButtonHeight(PropID)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get ButtonMenus() As TbrButtonMenus
Attribute ButtonMenus.VB_Description = "Returns a reference to a collection of button menu objects."
If PropButtonMenus Is Nothing Then
    Set PropButtonMenus = New TbrButtonMenus
    PropButtonMenus.FInit ShadowToolBar, PropID
End If
Set ButtonMenus = PropButtonMenus
End Property
